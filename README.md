/***************************************************************
 * FINANCEIRO FAMILIAR — (versão enxuta, consolidada)
 * Entrega: 1A (Core & Caches) + 1B (Listas, Cartões, DVs) + 1C (IDs/Status/Competências/Fechamento/Resumos/onEdit)
 * Notas rápidas:
 * - Centavos-first: use _partsCents_/_sumCents_ para qualquer cálculo com parcelas.
 * - EPS_CENT=0.005 reduz falso-positivo de arredondamento.
 * - CACHE_TTL=1800 diminui I/O em planilhas grandes.
 ***************************************************************/

/* ===================== [1A] Core & Caches ===================== */

/** Script properties */
function getCfg_(key, defv){ const p=PropertiesService.getScriptProperties(); const v=p.getProperty(key); return v!=null?v:defv; }

/** Ambiente (DEV/PROD, opcional) */
function env_(){ const e=(getCfg_('ENV','PROD')+'').toUpperCase(); return { DEV:e==='DEV', PROD:e!=='DEV' }; }

/** Spreadsheet alvo (por ID ou ativo) */
function SS_(){ const id=getCfg_('KEY_SPREADSHEET_ID',null); return id?SpreadsheetApp.openById(id):SpreadsheetApp.getActive(); }
/** ID “seguro” para cache/log/headless */
function _docIdSafe_(){ try{ const ss=SS_(); return ss && ss.getId ? ss.getId() : 'global'; }catch(_){ return 'global'; } }

/** Timezone + formatação (cache por planilha alvo) */
const __tz_cache__ = {};
function _tz_(){
  let ss;
  try { ss = SS_(); } catch(_) { ss = null; }
  const key = ss && ss.getId ? ss.getId() : 'global';
  if (__tz_cache__[key]) return __tz_cache__[key];
  let tz;
  try{
    tz = ss ? ss.getSpreadsheetTimeZone()
            : SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  }catch(_){
    tz = (Session.getScriptTimeZone && Session.getScriptTimeZone()) || 'America/Araguaina';
  }
  __tz_cache__[key] = tz || 'America/Araguaina';
  return __tz_cache__[key];
}
function formatDate_(d, fmt){ return Utilities.formatDate(new Date(d), _tz_(), fmt||'yyyy-MM-dd'); }

/** Moeda ⇄ centavos (robusto) */
function toCents_(v){
  if (v==null||v==='') return 0;
  if (typeof v==='number'){
    if(!isFinite(v)) throw new Error('Valor inválido: '+v);
    const s = /e/i.test(String(v)) ? v.toFixed(20) : String(v);
    return toCents_(s);
  }

  let s = String(v)
    .replace(/\u00A0/g,' ')
    .replace(/[−–—]/g,'-')
    .trim()
    .replace(/[R$\u20A8\u20AC\u00A3]/gi,'')
    .replace(/\s+/g,'');          

  let negative = false;
  if (/\(.*\)/.test(s)) { negative = true; s = s.replace(/[()]/g,''); }
  if (/-/.test(s))      { negative = true; s = s.replace(/-+/g,''); }

  if (s==='' || s==='.' || s===',') return 0;
  if (/^[.,]/.test(s)) s = '0'+s;
  if (/[.,]$/.test(s)) s += '0';

  if (s.indexOf(',')>=0 && s.indexOf('.')>=0){
    s = s.replace(/\.(?=\d{3}(\D|$))/g,'').replace(',', '.');
  } else if (s.indexOf(',')>=0){
    s = s.replace(',', '.');
  } else if (s.indexOf('.')>=0){
    const parts = s.split('.');
    const looksLikeThousands =
      parts.length > 1 &&
      parts[0] !== '0' &&
      /^\d{1,3}$/.test(parts[0]) &&
      parts.slice(1).every(seg => /^\d{3}$/.test(seg));
    if (looksLikeThousands){
      s = parts.join('');
    }
  }

  const m = s.match(/^(\d+)(?:\.(\d+))?$/);
  if(!m) throw new Error('Valor inválido: '+v);

  const intP = parseInt(m[1]||'0',10);
  const frac = m[2]||'';
  const d1 = frac[0] ? +frac[0] : 0;
  const d2 = frac[1] ? +frac[1] : 0;
  const d3 = frac[2] ? +frac[2] : 0;
  const cents = intP*100 + d1*10 + d2 + (d3>=5 ? 1 : 0);

  return (negative ? -1 : 1) * cents;
}
function fromCents_(c){ return (c||0)/100; }
function fmtBRL_(n){
  let x=(typeof n==='string')?Number(n.replace(',','.')):Number(n);
  if (!isFinite(x)) x = 0;
  if (Object.is(x,-0) || Math.abs(x) < 0.0005) x = 0;
  try{ return new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(x); }
  catch(_){ return 'R$ '+x.toFixed(2); }
}

/** LOG leve (aba LOGS) */
function _ensureLogsSheet_(){ const ss=SS_(); let sh=ss.getSheetByName('LOGS'); if(!sh) sh=ss.insertSheet('LOGS'); if(sh.getLastRow()<1) sh.appendRow(['Quando','Nível','ms','Obs']); return sh; }
function _log_(level, scope, msOrObj, obs){ try{ const sh=_ensureLogsSheet_(); const ms=(typeof msOrObj==='number')?msOrObj:''; const extra=(typeof msOrObj==='object'&&msOrObj)?' | '+JSON.stringify(msOrObj):''; sh.appendRow([new Date(),String(level||''),ms,`${scope||''}${obs?(' — '+obs):''}${extra}`]); }catch(_){ } }
const LOG_={ enabled:true, info:(...x)=>(LOG_.enabled&&Logger.log('[INFO] '+x.join(' '))), warn:(...x)=>(LOG_.enabled&&Logger.log('[WARN] '+x.join(' '))), error:(...x)=>(LOG_.enabled&&Logger.log('[ERROR] '+x.join(' '))) };

/** Ranges & header */
function effectiveRange_(sh, headerRows){ const hr=headerRows||1; const lr=Math.max(sh.getLastRow()-hr,0); const lc=sh.getLastColumn(); if(lr<=0||lc<=0) return null; return sh.getRange(hr+1,1,lr,lc); }
function headerMap_(sh, headerRow){ const r=sh.getRange(headerRow||1,1,1,sh.getLastColumn()); const names=r.getValues()[0].map(s=>(s||'').toString().trim()); const map={}; names.forEach((name,i)=>{ if(name) map[name]=i+1; }); return map; }

/** Lean mode (sem throttles) */
function lean_(){ return String(getCfg_('LEAN_MODE','true')).toLowerCase()==='true'; }

/** Flags & limiares */
const FEATURES={ NO_PROTECT:false, NO_COND_FORMAT:false };
const SAFETY  ={ MAX_PASTE_ROWS:2000, MAX_PASTE_COLS:20 };
const EPS_CENT=0.005;
const UTIL_THRESH={ WARN:0.30, ALERT:0.40 };
const ORC_ALERTS={ WARN_PCT:0.80, ALERT_PCT:1.00 };

/** Abas / Colunas */
const ABAS={ CONFIGURACOES:'Configurações', RECEITA:'Receita', DESPESAS_FIXAS_PREVISOES:'Despesas Fixas e Previsões', LANCAMENTO_DESPESA:'Lançamento de Despesa', PREVISAO_GASTOS:'Previsão de Gastos', CALENDARIO_FINANCEIRO:'Calendário Financeiro', RESUMO_ANUAL:'Resumo Anual', INDICADORES:'Indicadores', FATURAS_CARTAO:'Faturas de Cartão', RESUMO_FATURAS:'Resumo de Faturas', PARCELAS_CARTAO:'Parcelas do Cartão', JANEIRO:'Janeiro', FEVEREIRO:'Fevereiro', MARCO:'Março', ABRIL:'Abril', MAIO:'Maio', JUNHO:'Junho', JULHO:'Julho', AGOSTO:'Agosto', SETEMBRO:'Setembro', OUTUBRO:'Outubro', NOVEMBRO:'Novembro', DEZEMBRO:'Dezembro' };
const MESES=[ABAS.JANEIRO,ABAS.FEVEREIRO,ABAS.MARCO,ABAS.ABRIL,ABAS.MAIO,ABAS.JUNHO,ABAS.JULHO,ABAS.AGOSTO,ABAS.SETEMBRO,ABAS.OUTUBRO,ABAS.NOVEMBRO,ABAS.DEZEMBRO];
const COL={ DATA:2, SUBCATEGORIA:3, DETALHAMENTO:4, CATEGORIA:5, FORMA:6, PARCELAS:7, VALOR:8, VALOR_PARCELADO:9, STATUS:10, COMPETENCIA:11, CENTRO_CUSTO:12, TIPO:13, LIQUIDACAO:14, ID_EXTRATO:15 };
const COL2={ COMP_CONSUMO:16 };      // P
const COL_FP=17;                      // Q
const CARTOES_HEADER_ROW=17, CARTOES_FIRST_ROW=18, CARTOES_FIRST_COL=8, CARTOES_LAST_COL=13;
const LIMITE_LINHAS=9999;

/** Polyfill */
if (!Array.prototype.flat) Object.defineProperty(Array.prototype,'flat',{ value:function(d=1){ return this.reduce((a,v)=>a.concat(Array.isArray(v)?v.flat(d-1):v),[]); }});

/** Utils diversos */
function _stripDiacritics_(s){ try{ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,''); }catch(_){ return String(s||''); } }
function _normLower_(s){ return _stripDiacritics_(s).toLowerCase().trim(); }
function getSheetSmart_(preferido, alternates){
  const ss=SS_(); let sh=ss.getSheetByName(preferido); if(sh) return sh;
  for(const nm of (alternates||[])){ sh=ss.getSheetByName(nm); if(sh) return sh; }
  const alvo=_normLower_(preferido);
  for(const s of ss.getSheets()){ if(_normLower_(s.getName())===alvo) return s; }
  for(const s of ss.getSheets()){ if(_normLower_(s.getName()).indexOf(alvo)>=0) return s; }
  return null;
}
function _getCfg_(){ return getSheetSmart_(ABAS.CONFIGURACOES,['Configuracoes','Config','CFG']); }
function _getLanc_(){ return getSheetSmart_(ABAS.LANCAMENTO_DESPESA,['Lancamento de Despesa','Lançamentos','Lancamento']); }
function _norm(s){ try{ return String(s||'').toLowerCase().trim(); }catch(_){ return String(s||''); } }
function _fmtDate_(d, fmt='dd/MM/yyyy'){ return Utilities.formatDate(d, _tz_(), fmt); }
function _today_(){ const n=new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }
function _dateOnly_(d){ if(!(d instanceof Date)||isNaN(d)) return null; return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function _escRegex_(s){ return String(s||'').replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function _sanitizaDia_(v){ v=Number(v); return (v>=1&&v<=31)?v:null; }
function _isConciliadoStatus_(s){ return /\b(conciliad|pago|quitad|liquidad|baixad|compensad|ok)\b/i.test(String(s||'')); }
function _maybeToast_(msg){
  try{ SS_().toast(String(msg||'')); }
  catch(_){ try{ SpreadsheetApp.getActive().toast(String(msg||'')); }catch(__){ try{ Logger.log('[toast] '+msg); }catch(___){} } }
}
function _sameDate_(a,b){ const da=(a instanceof Date&&!isNaN(a))?a:null; const db=(b instanceof Date&&!isNaN(b))?b:null; if(da&&db) return da.getTime()===db.getTime(); return String(a||'')===String(b||''); }

/** Parse datas BR/ISO (estritas) */
function parseDateBR_(v){
  if(v instanceof Date && !isNaN(v)) return _dateOnly_(v);
  if(typeof v==='string'){
    const s=v.trim();
    let m=s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if(m){ const yyyy=+m[1], mm=+m[2], dd=+m[3]; const d=new Date(yyyy,mm-1,dd); if(d.getFullYear()===yyyy&&d.getMonth()===mm-1&&d.getDate()===dd) return _dateOnly_(d); return null; }
    m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if(m){ let dd=+m[1], mm=+m[2], yyyy=+m[3]; if(yyyy<100) yyyy+=(yyyy>=70?1900:2000); const d=new Date(yyyy,mm-1,dd); if(d.getFullYear()===yyyy&&d.getMonth()===mm-1&&d.getDate()===dd) return _dateOnly_(d); return null; }
  }
  return null;
}

/** Params/flags via Script Properties */
function param_(key, defv){ const v=getCfg_(key,null); return v!=null?v:defv; }
function features_(){
  const noFmtProp=param_('NO_COND_FORMAT',null); const noFmt=(noFmtProp!=null?(''+noFmtProp).toLowerCase()==='true':!!FEATURES.NO_COND_FORMAT);
  const noProtProp=param_('NO_PROTECT',null);     const noProt=(noProtProp!=null?(''+noProtProp).toLowerCase()==='true':!!FEATURES.NO_PROTECT);
  return { NO_COND_FORMAT:noFmt, NO_PROTECT:noProt };
}
function utilThresholds_(){ const t1=parseFloat(param_('UTIL_T1',UTIL_THRESH.WARN)); const t2=parseFloat(param_('UTIL_T2',UTIL_THRESH.ALERT)); return { WARN:t1, ALERT:t2 }; }

/** === DV helpers (só regrava quando muda) === */
function _dvSigList_(lista){ return 'LIST|'+(lista||[]).map(String).join('\u0001'); }
function _dvSigRule_(rule){
  try{
    if(!rule) return '';
    const crit=rule.getCriteriaType&&rule.getCriteriaType();
    const args=rule.getCriteriaValues&&rule.getCriteriaValues();
    if(crit===SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST){ const arr=(args&&args[0])?args[0]:[]; return _dvSigList_(arr); }
    if(crit===SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE){ const rng=(args&&args[0])?args[0]:null; if(rng) return 'RANGE|'+rng.getSheet().getSheetId()+'!'+rng.getA1Notation(); }
  }catch(_){}
  return '';
}
function _applyDVIfChanged_(range, lista, allowInvalid){
  if(!lista||!lista.length){ range.clearDataValidations(); return; }
  const wantSig=_dvSigList_(lista); let allSame=true; const existing=range.getDataValidations();
  for(let r=0; allSame && r<existing.length; r++){ const row=existing[r]; for(let c=0;c<row.length;c++){ if(_dvSigRule_(row[c])!==wantSig){ allSame=false; break; } } }
  if(allSame) return;
  const rule=SpreadsheetApp.newDataValidation().requireValueInList(lista,true).setAllowInvalid(!!allowInvalid).build();
  range.setDataValidation(rule);
}
function _applyDVRangeIfChanged_(range, targetRange, allowInvalid){
  if(!range||!targetRange){ try{ range && range.clearDataValidations(); }catch(_){ } return; }
  const wantSig='RANGE|'+targetRange.getSheet().getSheetId()+'!'+targetRange.getA1Notation();
  const existing=range.getDataValidations(); let allSame=existing&&existing.length>0;
  outer: for(let r=0; r<(existing?existing.length:0); r++){ const row=existing[r]||[]; for(let c=0;c<row.length;c++){ if(_dvSigRule_(row[c])!==wantSig){ allSame=false; break outer; } } }
  if(allSame) return;
  const rule=SpreadsheetApp.newDataValidation().requireValueInRange(targetRange,true).setAllowInvalid(!!allowInvalid).build();
  range.setDataValidation(rule);
}

/** === Cache leve (Config/Listas/Cartões) === */
const CACHE_TTL = 1800;   // 30 min
const SCRIPT_TTL= 21600;  // 6 h
function _cfgCacheKey_(s){ return 'cfg_cache_'+s; }
let _cardsMatchCache=null;

function _invalidateCfgCaches_(){
  const dc=CacheService.getDocumentCache();
  ['lists_v3','cards_v3','cfg_map_v2'].forEach(s=>{ try{ dc.remove(_cfgCacheKey_(s)); }catch(_){} });
  try{ const sc=CacheService.getScriptCache(); const ssId=_docIdSafe_(); ['lists_v3','cards_v3'].forEach(k=>sc.remove(ssId+'::'+k)); }catch(_){}
  _cardsMatchCache=null;
}
function _scGet_(k){ try{ return CacheService.getScriptCache().get(k); }catch(_){ return null; } }
function _scPut_(k,v,ttl){ try{ CacheService.getScriptCache().put(k,v,ttl); }catch(_){ } }
function ensureCfgCachesWarm_(cfgOpt){
  const ssId=_docIdSafe_(); const k1=ssId+'::lists_v3', k2=ssId+'::cards_v3';
  if(!_scGet_(k1)) try{ getCfgCached_(cfgOpt||_getCfg_(),true); }catch(_){}
  if(!_scGet_(k2)) try{ getCartoesCached_(cfgOpt||_getCfg_(),true); }catch(_){}
}

/** Centavos-first helpers (para 1B/1C usarem) */
function _partsCents_(valorTotalRaw, n, valorParcelaRaw){
  n = Math.max(1, parseInt(n,10) || 1);
  const totalCents = toCents_(valorTotalRaw);

  if (totalCents !== 0){
    const sign = totalCents < 0 ? -1 : 1;
    const abs  = Math.abs(totalCents);
    const base = Math.floor(abs / n), rest = abs - base*n;
    return Array.from({length:n}, (_, i) => sign * (base + (i < rest ? 1 : 0)));
  }

  const parc = toCents_(valorParcelaRaw);
  if (parc !== 0) return Array.from({length:n}, () => parc);
  return Array(n).fill(0);
}
function _sumCents_(arr){ return (arr||[]).reduce((a,b)=>a+(Number(b)||0),0); }

/** Auto-detecção de campos na aba de Config */
function _findLabelCell_(cfg, labels, searchCols, scanRows){
  const maxRows=Math.min(scanRows||50,cfg.getMaxRows()); const cols=searchCols&&searchCols.length?searchCols:[8,9,10]; const wanted=labels.map(_norm);
  for(const c of cols){ const vals=cfg.getRange(1,c,maxRows,1).getValues().flat(); for(let r=0;r<vals.length;r++){ const s=_norm(vals[r]); if(!s) continue; if(wanted.some(w=>s.includes(w))) return cfg.getRange(r+1,c); } }
  return null;
}
function _rightValueCell_(labelCell){
  if(!labelCell) return null; const sh=labelCell.getSheet(), row=labelCell.getRow(), lastCol=sh.getLastColumn();
  for(let c=labelCell.getColumn()+1; c<=Math.min(lastCol,labelCell.getColumn()+4); c++){ const cell=sh.getRange(row,c), v=cell.getValue(); if(!(v===""||v===null)) return cell; }
  return sh.getRange(row, Math.min(lastCol,labelCell.getColumn()+1));
}
function _cfgMap_(cfgIn){
  const cfg=cfgIn||_getCfg_(); if(!cfg) return { ANOREF_CELL:'I14', COMP_MODE_CELL:'I15', CARTOES_HEADER_ROW, CARTOES_FIRST_ROW, CARTOES_FIRST_COL, CARTOES_LAST_COL };
  const cache=CacheService.getDocumentCache(); const key=_cfgCacheKey_('cfg_map_v2'); const c=cache.get(key); if(c){ try{ return JSON.parse(c); }catch(_){ } }
  const lblAno=_findLabelCell_(cfg,['anoref','ano de referencia','ano de referência'],[8,9],60);
  const lblModo=_findLabelCell_(cfg,['modo','competencia','competência','caixa','consumo'],[8,9],60);
  const anoCell=_rightValueCell_(lblAno), modoCell=_rightValueCell_(lblModo);
  const map={ ANOREF_CELL: anoCell?anoCell.getA1Notation():'I14', COMP_MODE_CELL: modoCell?modoCell.getA1Notation():'I15', CARTOES_HEADER_ROW, CARTOES_FIRST_ROW, CARTOES_FIRST_COL, CARTOES_LAST_COL };
  try{ cache.put(key, JSON.stringify(map), CACHE_TTL); }catch(_){ }
  return map;
}
function _isEditedCell_(sheet,row,col,a1,editedRangeOpt){
  if(!a1) return false;
  const t=sheet.getRange(a1), r=editedRangeOpt||sheet.getRange(row,col,1,1);
  const r1=r.getRow(), c1=r.getColumn(), r2=r1+r.getNumRows()-1, c2=c1+r.getNumColumns()-1;
  const tr=t.getRow(), tc=t.getColumn(), tr2=tr+t.getNumRows()-1, tc2=tc+t.getNumColumns()-1;
  return !(r2<tr || r1>tr2 || c2<tc || c1>tc2);
}

/** Listas/Dados da aba Config (cacheados) */
function _cfgLastRow_(cfg){ const last=cfg.getLastRow(); const vals=last>2?cfg.getRange(3,2,Math.max(0,last-2),1).getValues().flat():[]; let end=vals.length; while(end>0&&String(vals[end-1]||'').trim()==='') end--; return Math.max(3,end+2); }
function _rng(cfg,c1,r0){ const last=_cfgLastRow_(cfg); return cfg.getRange(r0,c1,last-r0+1,1); }

function getCfgCached_(cfgIn, force){
  const cfg=cfgIn||_getCfg_(); if(!cfg) return { subs:[], cats:[], dets:[], subToCat:{}, catToSubs:{} };
  const dc=CacheService.getDocumentCache(); const ssId=_docIdSafe_(); const scKey=ssId+'::lists_v3';
  const dcKey=_cfgCacheKey_('lists_v3');
  if(!force){
    const sc=_scGet_(scKey); if(sc){ try{ return JSON.parse(sc); }catch(_){ } }
    const c=dc.get(dcKey); if(c){ try{ return JSON.parse(c); }catch(_){ } }
  }
  const last=_cfgLastRow_(cfg);
  const subs=_rng(cfg,2,3).getValues().slice(0,last-2).flat().filter(String);
  const cats=_rng(cfg,3,3).getValues().slice(0,last-2).flat().filter(String);
  const dets=_rng(cfg,5,3).getValues().slice(0,last-2).flat().filter(String);
  const subToCat={}; subs.forEach((s,i)=>{ const c=cats[i]; if(s&&c) subToCat[s]=c; });
  const catToSubs={}; subs.forEach((s,i)=>{ const c=cats[i]; if(c) (catToSubs[c]=(catToSubs[c]||[])).push(s); });
  const payload={ subs, cats, dets, subToCat, catToSubs };
  try{ dc.put(dcKey, JSON.stringify(payload), CACHE_TTL); }catch(_){ }
  _scPut_(scKey, JSON.stringify(payload), SCRIPT_TTL);
  return payload;
}

/** Leitura da tabela de cartões (H..M por padrão) */
function _rangeCartoes_(cfgIn){
  const cfg=cfgIn||_getCfg_(); if(!cfg) return [];
  const map=_cfgMap_(cfg); const sh=cfg; const lastRow=sh.getLastRow();
  if(lastRow<map.CARTOES_FIRST_ROW) return [];
  const nomes=sh.getRange(map.CARTOES_FIRST_ROW, map.CARTOES_FIRST_COL, lastRow-map.CARTOES_FIRST_ROW+1,1).getValues().flat();
  let endRow=map.CARTOES_FIRST_ROW-1;
  for(let i=0;i<nomes.length;i++){ const has=String(nomes[i]||'').trim()!==''; if(has) endRow=map.CARTOES_FIRST_ROW+i; }
  if(endRow<map.CARTOES_FIRST_ROW) return [];
  const numRows=endRow-map.CARTOES_FIRST_ROW+1;
  return sh.getRange(map.CARTOES_FIRST_ROW, map.CARTOES_FIRST_COL, numRows, map.CARTOES_LAST_COL-map.CARTOES_FIRST_COL+1).getValues()
    .filter(r=>String(r[0]||'').trim()!=='')
    .sort((a,b)=> String(b[0]||'').length - String(a[0]||'').length);
}

/** Cache leve para cartões (simétrico ao de listas) */
function getCartoesCached_(cfgIn, force){
  const cfg = cfgIn || _getCfg_(); 
  if (!cfg) return [];

  const dc   = CacheService.getDocumentCache();
  const ssId = _docIdSafe_();

  const scKey = ssId + '::cards_v3';
  const dcKey = _cfgCacheKey_('cards_v3');

  if (!force){
    const sc = _scGet_(scKey);
    if (sc){
      try { return JSON.parse(sc); } catch(_) { /* ignora cache corrompido */ }
    }
    const c = dc.get(dcKey);
    if (c){
      try { return JSON.parse(c); } catch(_) { /* ignora cache corrompido */ }
    }
  }

  const rows = _rangeCartoes_(cfg) || [];

  try { dc.put(dcKey, JSON.stringify(rows), CACHE_TTL); } catch(_){ /* sem stress */ }
  _scPut_(scKey, JSON.stringify(rows), SCRIPT_TTL);

  return rows;
}

/* ===================== [1B] Listas, Cartões, DVs & Ciclo ===================== */

/** EXACT/LOOSE para casar “Forma” com cartões */
function _cardsMatchMode_(){ const m=String(getCfg_('CARD_MATCH','LOOSE')||'').toUpperCase().trim(); return (m==='EXACT')?'EXACT':'LOOSE'; }

/** Compilação de regex p/ cartões (cacheada) — usa cache de cartões */
function _getCardsMatchers_(cfg){
  const c=cfg||_getCfg_(); const key=(c&&c.getSheetId?c.getSheetId():'default')+'::'+_cardsMatchMode_();
  if(_cardsMatchCache && _cardsMatchCache.key===key) return _cardsMatchCache.list;
  const rows=getCartoesCached_(c,false)||[]; 
  const mode=_cardsMatchMode_();
  const list=rows.filter(r=>String(r[0]||'').trim()).map(r=>{
    const nomeNorm=_stripDiacritics_(String(r[0]||'')).toLowerCase().trim();
    const esc=_escRegex_(nomeNorm);
    const re=(mode==='EXACT')? new RegExp(`^${esc}$`,'i') : new RegExp(`(?:^|[\\s\\-_/.,;()])${esc}(?:[\\s\\-_/.,;()]|$)`,'i');
    return { row:r, re };
  });
  _cardsMatchCache={ key, list };
  return list;
}
function ehCartao_(forma){ const f=_stripDiacritics_(String(forma||'')).toLowerCase().trim(); if(!f) return false; return _getCardsMatchers_(_getCfg_()).some(x=>x.re.test(f)); }
function _findCartaoRow_(forma){ const f=_stripDiacritics_(String(forma||'')).toLowerCase().trim(); if(!f) return null; const hit=_getCardsMatchers_(_getCfg_()).find(x=>x.re.test(f)); return hit?hit.row:null; }

/** Ciclo de cartão: devolve mês/ano da fatura e datas-chave */
function calcularCicloCartao(dataCompra, ini, fim, venc){
  ini=Number(ini); fim=Number(fim); venc=Number(venc);
  if(!(ini>=1&&ini<=31)||!(fim>=1&&fim<=31)||!(venc>=1&&venc<=31)){ const d=new Date(dataCompra); return { ano:d.getFullYear(), mes:d.getMonth()+1, primeiroDia:new Date(d.getFullYear(), d.getMonth(), 1) }; }
  function atMidday_(d){ const x=new Date(d); x.setHours(12,0,0,0); return x; }
  function safeDate_(y,m,d){ const max=new Date(y,m+1,0).getDate(); const dt=new Date(y,m,Math.min(d,max)); dt.setHours(12,0,0,0); return dt; }
  const p=atMidday_(dataCompra); const pDay=p.getDate();
  let closeMonthOffset; if(ini<=fim){ closeMonthOffset=(pDay>=fim)?1:0; } else { closeMonthOffset=(pDay>=ini)?1:0; if(pDay===fim) closeMonthOffset=1; }
  const closeBase=new Date(p.getFullYear(), p.getMonth()+closeMonthOffset,1);
  const dueMonthOffset=(venc>fim)?0:1; const dueFirstDay=new Date(closeBase.getFullYear(), closeBase.getMonth()+dueMonthOffset,1);
  const inicio=safeDate_(closeBase.getFullYear(), closeBase.getMonth() - (ini<=fim?0:1), ini);
  const fimDate=safeDate_(closeBase.getFullYear(), closeBase.getMonth(), fim);
  const vencDate=safeDate_(dueFirstDay.getFullYear(), dueFirstDay.getMonth(), venc);
  return { ano:dueFirstDay.getFullYear(), mes:dueFirstDay.getMonth()+1, primeiroDia:dueFirstDay, inicio, fim:fimDate, venc:vencDate };
}

/** Listas/DV (Categorias, Sub, Formas) */
function listaCategorias_(cfgIn){ const cfg=cfgIn||_getCfg_(); if(!cfg) return []; const last=_cfgLastRow_(cfg); return [...new Set(cfg.getRange(3,3,last-2,1).getValues().flat().filter(String))]; }
function listaSubcategorias_(cfgIn){ const cfg=cfgIn||_getCfg_(); if(!cfg) return []; const last=_cfgLastRow_(cfg); return cfg.getRange(3,2,last-2,1).getValues().flat().filter(String); }
function _cartaoTemCicloValido_(row){ const venc=_sanitizaDia_(row[1]), ini=_sanitizaDia_(row[2]), fim=_sanitizaDia_(row[3]); return !(venc==null||ini==null||fim==null); }

function listaFormasComCartoes_(cfgIn){
  const cfg=cfgIn||_getCfg_(); if(!cfg) return [];
  const vals=cfg.getRange(3,8,9,1).getValues().flat(); // H3:H11 (formas fixas)
  const fixas=vals.map(v=>String(v||'').trim()).filter(Boolean);
  const cartoes=getCartoesCached_(cfg,false)
    .map(r=>String(r[0]||'').trim())
    .filter(Boolean);
  const seen=new Set(), out=[]; [...fixas,...cartoes].forEach(v=>{ const k=_normLower_(v); if(k && !seen.has(k)){ seen.add(k); out.push(v.trim()); } });
  return out;
}
function atualizarMenuCategoriasLancamento(shIn, categorias){
  const sh=shIn||_getLanc_(); if(!sh) return; const firstRow=4;
  const lastUsed=Math.max(firstRow, sh.getLastRow()); const buffer=400;
  const totalRows=Math.min(LIMITE_LINHAS, (lastUsed-firstRow+1)+buffer); if(totalRows<=0) return;
  _applyDVIfChanged_(sh.getRange(firstRow, COL.CATEGORIA, totalRows,1), categorias, false);
}
function atualizarMenuSubcategoriasLancamento(shIn, subs){
  const sh=shIn||_getLanc_(); if(!sh) return; const firstRow=4;
  const lastUsed=Math.max(firstRow, sh.getLastRow()); const buffer=400;
  const totalRows=Math.min(LIMITE_LINHAS, (lastUsed-firstRow+1)+buffer); if(totalRows<=0) return;
  _applyDVIfChanged_(sh.getRange(firstRow, COL.SUBCATEGORIA, totalRows,1), subs, false);
}
function atualizarMenuFormasLancamento(shIn, formas){
  const sh=shIn||_getLanc_(); if(!sh) return; const firstRow=4;
  const lastUsed=Math.max(firstRow, sh.getLastRow()); const buffer=400;
  const totalRows=Math.min(LIMITE_LINHAS, (lastUsed-firstRow+1)+buffer); if(totalRows<=0) return;
  const rng=sh.getRange(firstRow, COL.FORMA, totalRows,1);
  if(!formas||!formas.length){ rng.clearDataValidations(); return; }
  _applyDVIfChanged_(rng, formas, false);
}

/** DVs numéricas + status (financeiro-safe) */
function ensureDVStatusLancamento_(){
  const sh=_getLanc_(); if(!sh) return; const firstRow=4;
  const lastUsed=Math.max(firstRow, sh.getLastRow()); const buffer=400;
  const totalRows=Math.min(LIMITE_LINHAS, (lastUsed-firstRow+1)+buffer); if(totalRows<=0) return;
  const lista=["Pendente","Conciliado"];
  const dvStatus=SpreadsheetApp.newDataValidation().requireValueInList(lista,true).setAllowInvalid(false).build();
  sh.getRange(firstRow, COL.STATUS, totalRows, 1).setDataValidation(dvStatus);
}
function ensureDVNumericasLancamento_(){
  const sh = _getLanc_(); if(!sh) return;
  const firstRow = 4;
  const lastUsed = Math.max(firstRow, sh.getLastRow());
  const buffer   = 400;
  const LIM      = Math.max(1, Number(typeof LIMITE_LINHAS!=='undefined' ? LIMITE_LINHAS : 2000) || 2000);
  const n        = Math.min(LIM, (lastUsed - firstRow + 1) + buffer);
  if (n<=0) return;

  const dvParc = SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(false).build();
  const dvData = SpreadsheetApp.newDataValidation().requireDateOnOrAfter(new Date(2000,0,1)).setAllowInvalid(false).build();
  const dvValor= SpreadsheetApp.newDataValidation().requireNumberBetween(-1e9, 1e9).setAllowInvalid(true).build();

  sh.getRange(firstRow, COL.PARCELAS,       n, 1).setDataValidation(dvParc);
  sh.getRange(firstRow, COL.DATA,           n, 1).setDataValidation(dvData);
  sh.getRange(firstRow, COL.VALOR,          n, 1).setDataValidation(dvValor);
  sh.getRange(firstRow, COL.VALOR_PARCELADO,n, 1).setDataValidation(dvValor);
}
/** DV dependente (Detalhamento por Sub) */
function atualizarMenusDinamicos(sheetIn, cfgIn, startRow, numRows){
  const sheet=sheetIn||_getLanc_(); const cfg=cfgIn||_getCfg_(); if(!sheet||!cfg||numRows<=0) return;
  const maxRow=_cfgLastRow_(cfg);
  const dados=cfg.getRange(3,5,maxRow-2,2).getValues(); // E=Detalhamento, F=Sub
  const mapa={}; for(let i=0;i<dados.length;i++){ const det=String(dados[i][0]||'').trim(); const sub=String(dados[i][1]||'').trim(); if(sub&&det) (mapa[sub]=(mapa[sub]||[])).push(det); }
  const subsEditadas=sheet.getRange(startRow, COL.SUBCATEGORIA, numRows,1).getValues().flat();
  const novas=subsEditadas.map(sub=>{ const opcoes=sub?(mapa[sub]||[]):null; if(opcoes&&opcoes.length){ return SpreadsheetApp.newDataValidation().requireValueInList(opcoes,true).setAllowInvalid(true).build(); } return null; });
  sheet.getRange(startRow, COL.DETALHAMENTO, numRows,1).setDataValidations(novas.map(v=>[v]));
}

/** Utilitário: corrigir DV da coluna FORMA sob demanda */
function corrigirDVFormaAgora_(){ const sh=_getLanc_(), cfg=_getCfg_(); if(!sh||!cfg) return; try{ const formas=listaFormasComCartoes_(cfg); atualizarMenuFormasLancamento(sh, formas); }catch(_){ } }

/* ===================== [1C] IDs, Status/Liquidação, Competências, Fechamento/Proteções,
 *      Resumos (mensal + cartões), Visual/CF, Recalcs e onEdit ===================== */

/* ========= Getters de ano/mode (I14/I15 com fallback I9/I10) ========= */
function getAnoRef_(cfgIn){
  const cfg = cfgIn || _getCfg_(); 
  if (!cfg) return (new Date()).getFullYear();

  function asYear(v){
    if (v instanceof Date && !isNaN(v)) return v.getFullYear();
    const s = String(v||'').trim();
    if (!s) return null;
    const d = new Date(s);
    if (!isNaN(d)) return d.getFullYear();
    const n = parseInt(s,10);
    return (isFinite(n) && n >= 1900 && n <= 9999) ? n : null;
  }

  try { const y = asYear(cfg.getRange('I14').getValue()); if (y) return y; } catch(_){}
  try { const y = asYear(cfg.getRange('I9').getValue());  if (y) return y; } catch(_){}
  return (new Date()).getFullYear();
}
function getCompMode_(cfgIn){
  const cfg = cfgIn || _getCfg_(); 
  if (!cfg) return 'CAIXA';
  const norm = s => String(s||'').toUpperCase().trim();

  let raw = null;
  try { raw = cfg.getRange('I15').getDisplayValue(); } catch(_){}
  if (!raw){ try { raw = cfg.getRange('I10').getDisplayValue(); } catch(_){} }

  raw = norm(raw);
  return raw.includes('CONSUMO') ? 'CONSUMO' : 'CAIXA';
}
function _metaMensalMap_(){
  const sh=getSheetSmart_(ABAS.PREVISAO_GASTOS,['Previsao de Gastos']); const mapa=new Map(); if(!sh) return mapa;
  const last=Math.max(2, sh.getLastRow()); if(last<2) return mapa;
  const itens=sh.getRange(2,1,last-1,1).getValues().flat();
  const metas=sh.getRange(2,3,last-1,1).getValues().flat();
  for(let i=0;i<itens.length;i++){ const k=String(itens[i]||'').trim(); const v=Number(metas[i]||0); if(k) mapa.set(k,v); }
  return mapa;
}
function _itemKeyFromSubDet_(sub, det){
  const cfg=_getCfg_(); if(!cfg) return String(sub||'').trim();
  const last=_cfgLastRow_(cfg);
  const dets=new Set(cfg.getRange(3,5,last-2,1).getValues().flat().filter(String));
  const d=String(det||'').trim(), s=String(sub||'').trim();
  return d && dets.has(d) ? d : s;
}
function _mesKey_(d){ const tz=_tz_(); const base=new Date(d.getFullYear(), d.getMonth(), 1); return Utilities.formatDate(base, tz, 'MM/yyyy'); }

/* ================= Batch writers (menos I/O, escritas em bloco) ================= */
function _setColValuesBatch_(sheet, col, entries){
  if(!entries||!entries.length) return; entries.sort((a,b)=>a.r-b.r);
  let i=0; while(i<entries.length){ let start=entries[i].r; let block=[[entries[i].v]]; let j=i+1;
    while(j<entries.length && entries[j].r===start+(j-i)){ block.push([entries[j].v]); j++; }
    sheet.getRange(start,col,block.length,1).setValues(block); i=j;
  }
}
function _setColFormatBatch_(sheet, col, entries){
  if(!entries||!entries.length) return; entries.sort((a,b)=>a.r-b.r);
  let i=0; while(i<entries.length){ let start=entries[i].r, fmt=entries[i].fmt, len=1, j=i+1;
    while(j<entries.length && entries[j].fmt===fmt && entries[j].r===start+(j-i)){ len++; j++; }
    sheet.getRange(start,col,len,1).setNumberFormat(fmt); i=j;
  }
}
function _setColNotesBatch_(sheet, col, entries){
  if(!entries||!entries.length) return; entries.sort((a,b)=>a.r-b.r);
  for(const {r, note} of entries){ sheet.getRange(r,col).setNote(note||''); }
}

/* ===================== IDs / Status / Competências defaults ===================== */
function _eqNum_(a,b){ return Math.abs(Number(a||0)-Number(b||0)) < EPS_CENT; }
function _parcelasExatas_(total, n){
  n = Math.max(1, parseInt(n,10)||1);
  const totC = toCents_(total);
  const sign = totC < 0 ? -1 : 1;
  const abs  = Math.abs(totC);
  const base = Math.floor(abs/n), rest = abs - base*n;
  return Array.from({length:n}, (_,i)=> sign*(base+(i<rest?1:0))/100);
}

function _gerarIdExtrato_(sheet, r, rowVals, forceOrOpts){
  const opts = (forceOrOpts && typeof forceOrOpts === 'object') ? forceOrOpts : { force: !!forceOrOpts, returnOnly: false };
  const force = !!opts.force, returnOnly = !!opts.returnOnly;

  const tz=_tz_();
  const d0=rowVals[0];
  const d=(d0 instanceof Date && !isNaN(d0)) ? _dateOnly_(d0) : parseDateBR_(d0);
  const dStr=d ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : "";

  const sub   = rowVals[COL.SUBCATEGORIA - COL.DATA] || "",
        det   = rowVals[COL.DETALHAMENTO - COL.DATA] || "",
        cat   = rowVals[COL.CATEGORIA    - COL.DATA] || "",
        forma = rowVals[COL.FORMA        - COL.DATA] || "";
  const nRaw  = Number(rowVals[COL.PARCELAS - COL.DATA] || 1);
  const n     = (isFinite(nRaw) && nRaw > 0) ? nRaw : 1;

  const totalCents = toCents_(rowVals[COL.VALOR - COL.DATA]);
  const parcCents  = toCents_(rowVals[COL.VALOR_PARCELADO - COL.DATA]);
  const centsBase  = (totalCents !== 0) ? totalCents : (parcCents * n);

  const payload=[ dStr, sub, det, cat, forma, n, (centsBase/100).toFixed(2) ].join("|");
  const hex=Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, payload).map(b=>('0'+(b&0xFF).toString(16)).slice(-2)).join('').toUpperCase();
  const id = "TX#" + hex.slice(0,10);

  if (!returnOnly && sheet && r){
    const cell = sheet.getRange(r, COL.ID_EXTRATO);
    if (!cell.getValue() || force) cell.setValue(id);
  }
  return id;
}
function _tratarMudancaStatus_(sheet, r, status){
  const liqCell=sheet.getRange(r, COL.LIQUIDACAO);
  if (_isConciliadoStatus_(status)){ if(!liqCell.getValue()) liqCell.setValue(_today_()); }
  else if (/pendente/i.test(String(status||""))){ liqCell.clearContent(); }
}
function _tiposDefault_(cfgIn){
  const cfg=cfgIn||_getCfg_(); if(!cfg) return 'Variável';
  const last=_cfgLastRow_(cfg);
  const lista=cfg.getRange(3,11,last-2,1).getValues().flat().filter(String);
  const pref=lista.find(t=>['variavel','variável'].includes(_norm(t)));
  return pref||lista[0]||'Variável';
}
function _centroCustoDefault_(cfgIn){
  const cfg=cfgIn||_getCfg_(); if(!cfg) return 'Casa';
  const last=_cfgLastRow_(cfg);
  const lista=cfg.getRange(3,10,last-2,1).getValues().flat().filter(String);
  const casa=lista.find(x=>_norm(x)==='casa'); return casa||lista[0]||'Casa';
}
function _rowFingerprint_(rowVals){
  const tz=_tz_(), d0=rowVals[0];
  const d=(d0 instanceof Date&&!isNaN(d0))?_dateOnly_(d0):parseDateBR_(d0);
  const dStr=d?Utilities.formatDate(d, tz, "yyyy-MM-dd"):"";

  const sub = rowVals[COL.SUBCATEGORIA - COL.DATA]||"",
        det = rowVals[COL.DETALHAMENTO - COL.DATA]||"",
        cat = rowVals[COL.CATEGORIA    - COL.DATA]||"",
        forma=rowVals[COL.FORMA        - COL.DATA]||"";
  const n   = Number(rowVals[COL.PARCELAS - COL.DATA]||1);

  const totalCents = toCents_(rowVals[COL.VALOR - COL.DATA]);
  const parcCents  = toCents_(rowVals[COL.VALOR_PARCELADO - COL.DATA]);
  const centsBase  = (totalCents !== 0) ? totalCents : (parcCents * (isFinite(n)&&n>0?n:1));

  const payload=[ dStr, sub, det, cat, forma, n, (centsBase/100).toFixed(2) ].join('|');
  const hex=Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,payload).map(b=>('0'+(b&0xFF).toString(16)).slice(-2)).join('');
  return hex.slice(0,10).toUpperCase();
}

/* ===================== Fechamento de mês & Proteções ===================== */
function _keyFechados_(){ return 'meses_fechados_v2'; }
function _competenciaDaLinha_(rowVals){
  const kVal=rowVals[COL.COMPETENCIA - COL.DATA];
  if(kVal instanceof Date && !isNaN(kVal)) return new Date(kVal.getFullYear(), kVal.getMonth(), 1);
  const k=String(kVal||'').trim();
  if(/^\d{2}\/\d{4}$/.test(k)){ const [mm,yyyy]=k.split('/').map(Number); return new Date(yyyy, mm-1, 1); }
  const d0=rowVals[0], d=(d0 instanceof Date && !isNaN(d0))?d0:parseDateBR_(d0);
  return d?new Date(d.getFullYear(), d.getMonth(), 1):null;
}
function _estaFechadoPorComp_(rowVals){
  const dp=PropertiesService.getDocumentProperties();
  const set=new Set((dp.getProperty(_keyFechados_())||'').split(',').filter(Boolean));
  const base=_competenciaDaLinha_(rowVals); if(!base) return false;
  const k=Utilities.formatDate(base, _tz_(), 'MM/yyyy'); return set.has(k);
}
function _addMesFechado_(mmYYYY){
  const dp=PropertiesService.getDocumentProperties();
  const cur=new Set((dp.getProperty(_keyFechados_())||'').split(',').filter(Boolean));
  cur.add(mmYYYY); dp.setProperty(_keyFechados_(), Array.from(cur).join(','));
}
function _rmMesFechado_(mmYYYY){
  const dp=PropertiesService.getDocumentProperties();
  const cur=new Set((dp.getProperty(_keyFechados_())||'').split(',').filter(Boolean));
  cur.delete(mmYYYY); dp.setProperty(_keyFechados_(), Array.from(cur).join(','));
}
function fecharMesAtual_(){
  const hoje=_today_();
  const k=Utilities.formatDate(new Date(hoje.getFullYear(),hoje.getMonth(),1), _tz_(), 'MM/yyyy');
  _addMesFechado_(k);
  try{ aplicarProtecaoLinhasFechadas_(); }catch(_){}
  _maybeToast_('🔒 Mês fechado: '+k);
}
function reabrirMesAtual_(){
  const hoje=_today_();
  const k=Utilities.formatDate(new Date(hoje.getFullYear(), hoje.getMonth(),1), _tz_(), 'MM/yyyy');
  _rmMesFechado_(k);
  try{ aplicarProtecaoLinhasFechadas_(); }catch(_){}
  _maybeToast_('🔓 Mês reaberto: '+k);
}
function _mergeContiguous(rows){
  rows=(rows||[]).slice().sort((a,b)=>a-b);
  const out=[]; let s=null,p=null;
  for(const r of rows){ if(s==null){ s=p=r; continue; } if(r===p+1){ p=r; } else { out.push([s,p]); s=p=r; } }
  if(s!=null) out.push([s,p]); return out;
}
function _colLetter_(n){ let s=''; for(; n>0; n=Math.floor((n-1)/26)){ s=String.fromCharCode(65+(n-1)%26)+s; } return s; }
function aplicarProtecaoLinhasFechadas_(){
  if (features_().NO_PROTECT) return;
  const sh=_getLanc_(); if(!sh) return;
  const PREFIX='LOCK::MFECHADO';

  (sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)||[])
    .filter(p=>(p.getDescription()||'').startsWith(PREFIX))
    .forEach(p=>{ try{ p.remove(); }catch(_){ } });

  const dp=PropertiesService.getDocumentProperties();
  const fech=(dp.getProperty(_keyFechados_())||'').split(',').filter(Boolean);
  if(!fech.length) return;

  const first=4, last=sh.getLastRow(); if(last<first) return;
  const lastCol=sh.getLastColumn(), tz=_tz_();

  const comps=sh.getRange(first, COL.COMPETENCIA, last-first+1,1).getValues().flat();
  const datas=sh.getRange(first, COL.DATA,        last-first+1,1).getValues().flat();
  const keyOf=(d)=>Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), 1), tz, 'MM/yyyy');

  const byMonth=new Map();
  for(let i=0;i<comps.length;i++){
    const base=(comps[i] instanceof Date && !isNaN(comps[i]))?comps[i]:(datas[i] instanceof Date && !isNaN(datas[i])?datas[i]:null);
    if(!base) continue;
    const key=keyOf(base); if(!fech.includes(key)) continue;
    const row=first+i; (byMonth.get(key)||byMonth.set(key,[]).get(key)).push(row);
  }

  const openCols=new Set([COL.STATUS, COL.CENTRO_CUSTO, COL.TIPO, COL.LIQUIDACAO]);
  function buildLockedSegments(fromCol, toCol){
    const segs=[]; let s=null;
    for(let c=fromCol;c<=toCol;c++){
      if(openCols.has(c)){ if(s!=null){ segs.push([s,c-1]); s=null; } }
      else { if(s==null) s=c; }
    }
    if(s!=null) segs.push([s,toCol]); return segs;
  }
  const colSegs=buildLockedSegments(COL.DATA, lastCol);

  let created=0;
  for(const [mmYYYY, rows] of byMonth.entries()){
    for(const [r0,r1] of _mergeContiguous(rows)){
      for(const [c1,c2] of colSegs){
        try{
          sh.getRange(r0,c1, r1-r0+1, c2-c1+1).protect()
            .setDescription(`${PREFIX} ${mmYYYY} ${_colLetter_(c1)}:${_colLetter_(c2)} ${r0}-${r1}`);
          created++;
        }catch(_){}
      }
    }
  }
  try{ _log_('INFO','aplicarProtecaoLinhasFechadas_',0,`protecoes=${created}`); }catch(_){}
}

/* ===================== Visual/CF da Previsão e Resumo ===================== */
function ensurePrevisaoProgressoVisual_(){
  const sh=getSheetSmart_(ABAS.PREVISAO_GASTOS,['Previsao de Gastos']); if(!sh) return;
  const lastRow=Math.max(2, sh.getLastRow()); if(lastRow<2) return;
  sh.getRange(1,12,1,2).setValues([['charttype','bar']]); // L1:M1
  sh.getRange(2,12,1,2).setValues([['max',1]]);
  sh.getRange(1,14,1,2).setValues([['warn',0.8]]);        // N1:O1
  sh.getRange(2,14,1,2).setValues([['alert',1]]);
  try{ sh.hideColumns(12,4); }catch(_){}
  let loc='';
  try{ loc=(SS_().getSpreadsheetLocale()||'').toLowerCase(); }
  catch(_){ try{ loc=(SpreadsheetApp.getActive().getSpreadsheetLocale()||'').toLowerCase(); }catch(__){ loc=''; } }
  const SEP=(/^(pt|fr|de|es)/.test(loc))?';':',';
  sh.getRange(2,9,lastRow-1,1).setFormulaR1C1(`=IFERROR(MAX(0${SEP}-RC[-3]))`); // I
  sh.getRange(2,10,lastRow-1,1).setFormulaR1C1(`=MAX(0${SEP}MIN(1${SEP}1-RC[-1]))`); // J
  sh.getRange(2,11,lastRow-1,1).setFormulaR1C1(`=IF(RC[-1]>=R2C15${SEP}"Estourou"${SEP}IF(RC[-1]>=R1C15${SEP}"Atenção"${SEP}"OK"))`); // K
  sh.getRange(2,8,lastRow-1,1).setFormulaR1C1(`=IFERROR(SPARKLINE(RC[1]${SEP}R1C12:R2C13))`); // H
  try{ sh.getRange(2,9,lastRow-1,2).setNumberFormat('0.00%'); }catch(_){ }
}
function ensureCondFormatPrevisao_(){
  if (features_().NO_COND_FORMAT) return;
  const sh=getSheetSmart_(ABAS.PREVISAO_GASTOS,['Previsao de Gastos']); if(!sh) return;
  const { WARN_PCT:warn, ALERT_PCT:alert }=ORC_ALERTS;
  try{
    const rules=sh.getConditionalFormatRules(); const keep=[];
    const touchesHJ=(g)=>{ if(g.getSheet().getSheetId()!==sh.getSheetId()) return false;
      const c=g.getColumn(), w=g.getNumColumns(), r=g.getRow(), h=g.getNumRows();
      const lastC=c+w-1, lastR=r+h-1; const hitH=(8>=c&&8<=lastC)&&(lastR>=2); const hitJ=(10>=c&&10<=lastC)&&(lastR>=2); return hitH||hitJ; };
    for(const r of rules){ const rr=(r.getRanges&&r.getRanges())||[]; if(!rr.some(touchesHJ)) keep.push(r); }
    const ours=[
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(alert).setBackground('#F8D7DA').setRanges([sh.getRange('J2:J')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(warn,alert).setBackground('#FFF3CD').setRanges([sh.getRange('J2:J')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(warn).setBackground('#D4EDDA').setRanges([sh.getRange('J2:J')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$J2>='+alert).setBackground('#F8D7DA').setRanges([sh.getRange('H2:H')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($J2>='+warn+',$J2<'+alert+')').setBackground('#FFF3CD').setRanges([sh.getRange('H2:H')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$J2<'+warn).setBackground('#D4EDDA').setRanges([sh.getRange('H2:H')]).build()
    ];
    sh.setConditionalFormatRules(keep.concat(ours));
  }catch(_){}
}
function ensureCondFormatResumoUtil_(){
  if (features_().NO_COND_FORMAT) return;
  const sh=getSheetSmart_(ABAS.RESUMO_FATURAS,['Resumo de Faturas','Resumo']); if(!sh) return;
  const { WARN:t1, ALERT:t2 }=utilThresholds_();
  try{
    const rules=sh.getConditionalFormatRules(), keep=[];
    for(const r of rules){
      const ranges=(r.getRanges&&r.getRanges())||[];
      const isH2H=ranges.length>0 && ranges.every(g=> g.getSheet().getSheetId()===sh.getSheetId() && g.getColumn()===8 && g.getNumColumns()===1 && g.getRow()>=2 );
      if(!isH2H) keep.push(r);
    }
    const ours=[
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(t2).setBackground('#F8D7DA').setRanges([sh.getRange('H2:H')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(t1,t2).setBackground('#FFF3CD').setRanges([sh.getRange('H2:H')]).build()
    ];
    sh.setConditionalFormatRules(keep.concat(ours));
  }catch(_){}
}

/* ===================== Gasto por item (centavos-first) ===================== */
function _gastoNoMesParaItem_(refMesKey, itemKey, modo /* 'both'|'real'|'prev' */='both', opts){
  opts = opts || {};
  const paidKeys   = opts.paidKeys || null;              
  const usarResumo = !!opts.considerarResumoComoReal;    

  const lanc = _getLanc_(), cfg = _getCfg_();
  if (!lanc || !cfg || !refMesKey || !itemKey) return 0;

  const compMode = getCompMode_(cfg);
  const last = lanc.getLastRow();
  if (last < 4) return 0;

  const rows      = lanc.getRange(4, COL.DATA, last-3, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues();
  const statuses  = lanc.getRange(4, COL.STATUS,     last-3, 1).getValues().flat();
  const liqs      = lanc.getRange(4, COL.LIQUIDACAO, last-3, 1).getValues().flat();

  const lastCfg   = _cfgLastRow_(cfg);
  const detOfRaw  = (lastCfg>2) ? cfg.getRange(3,5,lastCfg-2,1).getValues().flat().filter(Boolean) : [];
  const detsOficiaisNorm = new Set(detOfRaw.map(x => _normLower_(_stripDiacritics_(String(x)))));

  const tz   = _tz_(), hoje = _today_();
  const kMes = (d) => Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), 1), tz, 'MM/yyyy');
  const want = _normLower_(_stripDiacritics_(String(itemKey||'')));

  let somaCents = 0;

  for (let i=0; i<rows.length; i++){
    const r            = rows[i];
    const statusLinha  = String(statuses[i]||'').trim();
    const liqRaw       = liqs[i];

    const [dataRaw, sub, det, /*cat*/, forma, parcelasRaw, valorRaw, valorParcRaw] = r;

    const detStr = String(det||'').trim();
    const subStr = String(sub||'').trim();
    const detN   = _normLower_(_stripDiacritics_(detStr));
    const subN   = _normLower_(_stripDiacritics_(subStr));

    const detEhOficial = detStr && detsOficiaisNorm.has(detN);
    const hit = detEhOficial ? (detN === want) : (detN === want || subN === want);
    if (!hit) continue;

    const dt = (dataRaw instanceof Date && !isNaN(dataRaw)) ? dataRaw : parseDateBR_(dataRaw);
    if (!dt) continue;

    const n          = parseInt(parcelasRaw,10) || 1;
    const partsCents = _partsCents_(valorRaw, n, valorParcRaw);

    if (getCompMode_(cfg) === 'CONSUMO'){
      if (kMes(dt) !== refMesKey) continue;
      const soma   = _sumCents_(partsCents);
      const isReal = _dateOnly_(dt).getTime() <= _dateOnly_(hoje).getTime();
      if (modo==='both' || (modo==='real'&&isReal) || (modo==='prev'&&!isReal)) somaCents += soma;
      continue;
    }

    const cartRow = _findCartaoRow_(forma);

    if (!cartRow){
      if (kMes(dt) !== refMesKey) continue;
      const soma   = _sumCents_(partsCents);
      const liqOk  = (liqRaw instanceof Date && !isNaN(liqRaw)) || !!parseDateBR_(liqRaw);
      const concOk = _isConciliadoStatus_(statusLinha);
      const isReal = (concOk || liqOk) && (_dateOnly_(dt).getTime() <= _dateOnly_(hoje).getTime());
      if (modo==='both' || (modo==='real'&&isReal) || (modo==='prev'&&!isReal)) somaCents += soma;
      continue;
    }

    const venc = _sanitizaDia_(cartRow[1]), ini = _sanitizaDia_(cartRow[2]), fim = _sanitizaDia_(cartRow[3]);

    for (let p=0; p<partsCents.length; p++){
      const v = partsCents[p] || 0;
      if (v === 0) continue;

      const dParc = new Date(dt);
      dParc.setMonth(dParc.getMonth()+p);

      let mesKey = kMes(dParc);
      let isReal = false;

      if (venc != null && ini != null && fim != null){
        const { mes, ano } = calcularCicloCartao(dParc, ini, fim, venc);
        mesKey = kMes(new Date(ano, mes-1, 1));

        if (usarResumo && paidKeys){
          const keyFat = (_stripDiacritics_(String(cartRow[0]||'')).toUpperCase()) + '||' + mesKey;
          if (paidKeys.has(keyFat)) isReal = true;
        }
      } else {
        isReal = false;
      }

      if (mesKey !== refMesKey) continue;
      if (modo==='both' || (modo==='real'&&isReal) || (modo==='prev'&&!isReal)) somaCents += v;
    }
  }

  return _r2(fromCents_(somaCents));
}

/* ===================== Resumo mensal (com cartões) ===================== */
function atualizarResultadosMensaisComCartoes(shLanc, shMes, shCfg, mesNum /*1-12*/, startRow, endRow){
  if(!shLanc||!shMes||!shCfg) return;
  mesNum=Math.max(1, Math.min(12, parseInt(mesNum,10)||1));

  const tz=_tz_(), hoje=_today_();
  const anoRef=getAnoRef_(shCfg)||(new Date()).getFullYear();
  const compMode=getCompMode_(shCfg);

  const paidKeys = _getPaidKeysResumo_();

  const somaRealPorSubCents=new Map();
  const somaPrevPorSubCents=new Map();
  const addC=(map, sub, cents)=>{ const k=String(sub||'').trim(); if(!k) return; const v=Number(cents)||0; if(!Number.isFinite(v)) return; map.set(k,(map.get(k)||0)+v); };

  const lastLanc=shLanc.getLastRow();
  if(lastLanc>=4){
    const rows=shLanc.getRange(4, COL.DATA, lastLanc-3, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues();
    const statuses=shLanc.getRange(4, COL.STATUS, lastLanc-3, 1).getValues().flat();

    for(let i=0;i<rows.length;i++){
      const [data, subRaw, , , forma, parcelasRaw, valorRaw, valorParcRaw]=rows[i];
      const statusLinha=String(statuses[i]||"").trim();
      const sub=String(subRaw||'').trim(); if(!sub) continue;

      const n=parseInt(parcelasRaw,10)||1;
      const partsCents=_partsCents_(valorRaw, n, valorParcRaw);

      const dt=(data instanceof Date && !isNaN(data))?data:parseDateBR_(data); if(!(dt instanceof Date)||isNaN(dt)) continue;

      if(compMode==='CONSUMO'){
        const soma=_sumCents_(partsCents);
        const m=dt.getMonth()+1, a=dt.getFullYear();
        if(a===anoRef && m===mesNum){
          const futuro=_dateOnly_(dt).getTime()>_dateOnly_(hoje).getTime();
          addC(futuro?somaPrevPorSubCents:somaRealPorSubCents, sub, soma);
        }
        continue;
      }

      const cartRow=_findCartaoRow_(forma);

      if(!cartRow){
        const soma=_sumCents_(partsCents);
        const m=dt.getMonth()+1, a=dt.getFullYear();
        if(a===anoRef && m===mesNum){
          const jaPassou=_dateOnly_(dt).getTime()<=_dateOnly_(hoje).getTime();
          if(!jaPassou) addC(somaPrevPorSubCents, sub, soma);
          else if(_isConciliadoStatus_(statusLinha)) addC(somaRealPorSubCents, sub, soma);
          else addC(somaPrevPorSubCents, sub, soma);
        }
        continue;
      }

      const cartaoNome=String(cartRow[0]||"").trim();
      const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]);

      if(venc==null||ini==null||fim==null){
        const soma=_sumCents_(partsCents);
        const m=dt.getMonth()+1, a=dt.getFullYear();
        if(a===anoRef && m===mesNum){
          addC(somaPrevPorSubCents, sub, soma);
        }
        continue;
      }

      for(let p=0;p<partsCents.length;p++){
        const dataParc=new Date(dt); dataParc.setMonth(dataParc.getMonth()+p);
        const { mes:mFat, ano:aFat }=calcularCicloCartao(dataParc, ini, fim, venc);
        if(aFat!==anoRef || mFat!==mesNum) continue;

        const vParcCents=partsCents[p]||0; if(vParcCents===0) continue;

        const keyBase=_stripDiacritics_(cartaoNome).toUpperCase();
        const kFat=keyBase+'||'+Utilities.formatDate(new Date(aFat, mFat-1, 1), tz, "MM/yyyy");
        const pagaNoResumo=paidKeys.has(kFat);

        if(pagaNoResumo){
          addC(somaRealPorSubCents, sub, vParcCents);
        }else{
          addC(somaPrevPorSubCents, sub, vParcCents);
        }
      }
    }
  }

  const colSub=[2,6,10,14,18,22,26,30,34];
  const colRes=[3,7,11,15,19,23,27,31,35];
  const colPrev=[4,8,12,16,20,24,28,32,36];

  const linhas=Math.max(0, endRow-startRow+1); if(linhas<=0) return;

  const escrever=(cols, getterCents)=>{
    for(let k=0;k<cols.length;k++){
      const cSub=colSub[k], cOut=cols[k];
      const labels=shMes.getRange(startRow,cSub,linhas,1).getValues().flat();
      const out=labels.map(lbl=>{
        const key=String(lbl||'').trim(); const cents=key?getterCents(key):0;
        const val=_r2(fromCents_(cents||0)); return [Number.isFinite(val)?val:0];
      });
      shMes.getRange(startRow,cOut,linhas,1).setValues(out);
    }
  };
  escrever(colRes, key=>somaRealPorSubCents.get(key)||0);
  escrever(colPrev, key=>somaPrevPorSubCents.get(key)||0);
}

/* ===================== Alertas de orçamento por linha ===================== */
function _r2(n){ if(!isFinite(n)) throw new Error("Valor inválido: "+n); return Math.round(n*100)/100; }
function _orcamentoAvisarSeUltrapassarLinha_(sheet, r){
  const cfg=_getCfg_(); if(!cfg) return;

  const row=sheet.getRange(r, COL.DATA, 1, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues()[0];
  const dataRaw=row[0], sub=row[1], det=row[2], forma=row[4];
  const dt=(dataRaw instanceof Date && !isNaN(dataRaw))?dataRaw:parseDateBR_(dataRaw); if(!dt) return;

  const compMode=getCompMode_(cfg);
  let refDate=new Date(dt.getFullYear(), dt.getMonth(), 1);
  if(compMode==='CAIXA'){
    const cartRow=_findCartaoRow_(forma);
    if(cartRow){
      const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]);
      if (venc!=null && ini!=null && fim!=null){
        const info=calcularCicloCartao(dt, ini, fim, venc);
        refDate=new Date(info.ano, info.mes-1, 1);
      }
    }
  }
  const refKey=_mesKey_(refDate);
  const itemKey=_itemKeyFromSubDet_(sub, det); if(!itemKey) return;

  const metas=_metaMensalMap_(); const meta=Number(metas.get(itemKey)||0); if(!(meta>0)) return;

  const paidKeys = _getPaidKeysResumo_();
  const gastoMes=_gastoNoMesParaItem_(refKey, itemKey, 'both', { paidKeys, considerarResumoComoReal:true });

  const ratio=gastoMes/meta;
  if(ratio>=ORC_ALERTS.ALERT_PCT){
    _maybeToast_(`🚨 Orçamento: "${itemKey}" em ${refKey} estourou (R$ ${gastoMes.toFixed(2)} / R$ ${meta.toFixed(2)})`);
  }else if(ratio>=ORC_ALERTS.WARN_PCT){
    _maybeToast_(`⚠️ Orçamento: "${itemKey}" em ${refKey} atingiu ${Math.round(ratio*100)}% da meta (R$ ${gastoMes.toFixed(2)} / R$ ${Number(meta).toFixed(2)})`);
  }
}

/* ===================== Recalcular agora (lean) ===================== */
function withDocLock_(fn, timeoutMs){
  const lock = LockService.getDocumentLock();
  const t = Math.max(1, Number(timeoutMs)||20000);
  if (!lock.tryLock(t)){ _maybeToast_('⏳ Operação em andamento. Tente novamente em instantes.'); return; }
  try{ return fn && fn(); } finally { try{ lock.releaseLock(); }catch(_){ } }
}
function doRecalcsNow_(opts={}){
  return withDocLock_(()=>{
    try{
      if(opts.faturas){
        try{ if(typeof gerarFaturasDeCartao_==='function') gerarFaturasDeCartao_(); else _log_('WARN','doRecalcsNow_',0,'missing gerarFaturasDeCartao_'); }catch(err){ _log_('ERROR','doRecalcsNow_',0,'gerarFaturasDeCartao_: '+(err&&err.message||err)); }
        try{ if(typeof atualizarResumoFaturas_==='function') atualizarResumoFaturas_(); else _log_('WARN','doRecalcsNow_',0,'missing atualizarResumoFaturas_'); }catch(err){ _log_('ERROR','doRecalcsNow_',0,'atualizarResumoFaturas_: '+(err&&err.message||err)); }
        try{ if(typeof rebuildParcelasCartao_==='function') rebuildParcelasCartao_(); else _log_('WARN','doRecalcsNow_',0,'missing rebuildParcelasCartao_'); }catch(err){ _log_('ERROR','doRecalcsNow_',0,'rebuildParcelasCartao_: '+(err&&err.message||err)); }
        try{ if(typeof sincronizarStatusLancamentosComResumo_==='function') sincronizarStatusLancamentosComResumo_({ downgrade:true }); else _log_('WARN','doRecalcsNow_',0,'missing sincronizarStatusLancamentosComResumo_'); }catch(err){ _log_('ERROR','doRecalcsNow_',0,'sincronizarStatusLancamentosComResumo_: '+(err&&err.message||err)); }
      }
    }catch(_){}

    try{
      if(typeof atualizarPrevisaoCompleta_==='function') { atualizarPrevisaoCompleta_({ overwriteMetas:false }); }
      else { _log_('WARN','doRecalcsNow_',0,'missing atualizarPrevisaoCompleta_'); }
      ensurePrevisaoProgressoVisual_();
      ensureCondFormatPrevisao_();
      runBudgetSweep_();
    }catch(err){ _log_('ERROR','doRecalcsNow_',0,'previsao/budget: '+(err&&err.message||err)); }

    try{ ensureCondFormatResumoUtil_(); }catch(_){}
  }, 20000);
}

/* ===================== Sweep de orçamento (toast compacto) ===================== */
function runBudgetSweep_(){
  const mapa=_metaMensalMap_(); if(!mapa.size) return;
  const hoje=_today_(), refKey=Utilities.formatDate(new Date(hoje.getFullYear(), hoje.getMonth(), 1), _tz_(), 'MM/yyyy');

  const paidKeys = _getPaidKeysResumo_();

  const hits=[];
  for(const [item, meta] of mapa.entries()){
    const gasto=_gastoNoMesParaItem_(refKey, item, 'both', { paidKeys, considerarResumoComoReal:true });
    const pct=meta?gasto/meta:0;
    if(pct>=ORC_ALERTS.WARN_PCT){
      const tag=(pct>=ORC_ALERTS.ALERT_PCT?'🚨':'⚠️');
      hits.push(`${tag} ${item}: ${Math.round(pct*100)}% (R$ ${gasto.toFixed(2)} / R$ ${Number(meta).toFixed(2)})`);
    }
  }
  if(hits.length){
    const msg=hits.length<=3?hits.join('\n'):`${hits.slice(0,3).join('\n')}\n… +${hits.length-3} itens`;
    _maybeToast_(msg);
  }
}

function _getPaidKeysResumo_(){
  const set = new Set();
  const shRes = getSheetSmart_(ABAS.RESUMO_FATURAS, ['Resumo de Faturas','Resumo']);
  if (!shRes) return set;

  const tz = _tz_();
  const last = shRes.getLastRow();
  if (last < 2) return set;

  const rows = shRes.getRange(2, 1, last-1, 9).getValues();
  for (const r of rows){
    const cartao = String(r[0]||'').trim();
    const mesVal = r[1];
    const pendStr = String(r[3]??'').trim();
    if (!(cartao && mesVal)) continue;

    const mesDate = (mesVal instanceof Date && !isNaN(mesVal)) ? new Date(mesVal.getFullYear(), mesVal.getMonth(), 1) : null;
    const mesStr  = mesDate ? Utilities.formatDate(new Date(mesDate.getFullYear(), mesDate.getMonth(), 1), tz, 'MM/yyyy')
                            : String(mesVal||'').trim();

    const key = (_stripDiacritics_(cartao).toUpperCase()) + '||' + mesStr;
    const pendOk = Math.abs(toCents_(pendStr)/100) <= EPS_CENT;
    if (pendOk) set.add(key);
  }
  return set;
}

/* ===================== onEdit (robusto para nomes de abas variantes) ===================== */
function onEdit(e){
  try{
    if(!e || !e.range) return;

    const { range } = e;
    const sheet = range.getSheet();
    const nomeAba = sheet.getName();
    const nomeAbaNorm = _normLower_(nomeAba);

    const colIni = range.getColumn(), rowIni = range.getRow();
    const numRows = range.getNumRows(), numCols = range.getNumColumns();
    const colFim = colIni + numCols - 1;
    const colHit = (...cols) => cols.some(c => c >= colIni && c <= colFim);
    const singleCellEdit = (numRows === 1 && numCols === 1);

    if (numRows > (SAFETY.MAX_PASTE_ROWS||2000) || numCols > (SAFETY.MAX_PASTE_COLS||20)){
      _maybeToast_(`⚠️ Colagem grande detectada (${numRows}×${numCols}). Quebre em partes ou use "Recalcular agora".`);
      try{ _log_('WARN','onEdit_guard',{rows:numRows,cols:numCols}); }catch(_){}
      return;
    }

    const oldV = (typeof e.oldValue==='undefined') ? '' : e.oldValue;
    const newV = (typeof e.value   ==='undefined') ? '' : e.value;
    try{
      if (singleCellEdit){
        const isCfg = [ABAS.CONFIGURACOES,'Configurações','Configuracoes','Config','CFG'].filter(Boolean).map(n=>_normLower_(String(n))).includes(nomeAbaNorm);
        if (isCfg){ _log_('AUDIT','CONFIG',0,`${sheet.getName()}!${range.getA1Notation()} | ${oldV} → ${newV}`); }
      }
    }catch(_){}

    const shCfg = _getCfg_(), shLanc = _getLanc_();
    if(!shCfg || !shLanc) return;

    const cartRowCache = new Map();
    function getCartRow(forma){
      const key=_normLower_(String(forma||'')); if(!key) return null;
      if(cartRowCache.has(key)) return cartRowCache.get(key);
      let r=null; try{ r=_findCartaoRow_(forma); }catch(_){ r=null; }
      cartRowCache.set(key,r); return r;
    }

    const cardCentersByName = new Map();
    try{
      const rows = getCartoesCached_(shCfg,false) || [];
      for (const rr of rows){
        const nome = String(rr[0]||'').trim(); if(!nome) continue;
        const cc   = String(rr[5]||'').trim();
        const key  = _normLower_(_stripDiacritics_(nome));
        if (cc && !cardCentersByName.has(key)) cardCentersByName.set(key, cc);
      }
    }catch(_){}

    /* ========== FATURAS DE CARTÃO ========== */
    if (nomeAbaNorm === _normLower_(ABAS.FATURAS_CARTAO)){
      const tocou = [5,6,8,9,10].some(c => c>=colIni && c<=colFim);
      if (tocou){ doRecalcsNow_({ faturas:true, resumo:true }); }
      return;
    }

    /* ========== CONFIGURAÇÕES ========== */
    if ([ABAS.CONFIGURACOES,'Configurações','Configuracoes','Config','CFG'].filter(Boolean).map(n=>_normLower_(String(n))).includes(nomeAbaNorm)){
      _invalidateCfgCaches_();

      const map         = _cfgMap_(shCfg);
      const isAnoRef    = _isEditedCell_(sheet,rowIni,colIni,map.ANOREF_CELL,range);
      const isCompMode  = _isEditedCell_(sheet,rowIni,colIni,map.COMP_MODE_CELL,range);
      const colIsCards  = (colIni>=map.CARTOES_FIRST_COL && colIni<=map.CARTOES_LAST_COL && rowIni>=map.CARTOES_FIRST_ROW);
      const linhaFim    = rowIni + numRows - 1;
      const isFormasIdxEdit = (colIni<=8 && colFim>=8 && rowIni<=11 && linhaFim>=3);

      const mudou = (('value' in e)||('oldValue' in e)) ? (e.value!==e.oldValue) : true;
      if(!mudou) return;

      if (colIsCards || isAnoRef || isFormasIdxEdit){
        try{ atualizarMenuFormasLancamento(shLanc, listaFormasComCartoes_(shCfg)); }catch(_){}
        try{ aplicarProtecaoLinhasFechadas_(); }catch(_){}
        doRecalcsNow_({ faturas:true, previsao:true, resumo:true });
        return;
      }
      if (isCompMode){
        doRecalcsNow_({ faturas:true, previsao:true, resumo:true });
        return;
      }
      if (colHit(2,3,5)){
        const cats = listaCategorias_(shCfg), subs = listaSubcategorias_(shCfg);
        atualizarMenuCategoriasLancamento(shLanc, cats);
        atualizarMenuSubcategoriasLancamento(shLanc, subs);
        try{ atualizarMenuFormasLancamento(shLanc, listaFormasComCartoes_(shCfg)); }catch(_){}
        doRecalcsNow_({ previsao:true, resumo:true });
        return;
      }
      return;
    }

    /* ========== LANÇAMENTO DE DESPESA ========== */
    const isLanc = [ABAS.LANCAMENTO_DESPESA,'Lançamento de Despesa','Lancamento de Despesa','Lançamentos','Lancamento']
      .filter(Boolean).map(n=>_normLower_(String(n))).includes(nomeAbaNorm);

    if (isLanc){
      if (rowIni < 4) return;

      const relevante = colHit(
        COL.DATA, COL.SUBCATEGORIA, COL.DETALHAMENTO, COL.CATEGORIA, COL.FORMA,
        COL.PARCELAS, COL.VALOR, COL.VALOR_PARCELADO, COL.STATUS, COL.LIQUIDACAO
      );
      if (!relevante) return;

      const cfgLists      = getCfgCached_(shCfg);
      const subsCfg       = cfgLists.subs;
      const catsCfg       = cfgLists.cats;
      const mapSubToCat   = cfgLists.subToCat;
      const tiposDefault  = _tiposDefault_(shCfg);
      const centroDefault = _centroCustoDefault_(shCfg);
      const categoriasLista = [...new Set(catsCfg)];

      let impactaPrev    = colHit(COL.DATA, COL.SUBCATEGORIA, COL.DETALHAMENTO, COL.FORMA, COL.PARCELAS, COL.VALOR, COL.VALOR_PARCELADO, COL.CATEGORIA);
      let impactaFaturas = colHit(COL.FORMA, COL.PARCELAS, COL.DATA, COL.VALOR, COL.VALOR_PARCELADO);
      const precisaAtualizarMenusFixos = colHit(COL.CATEGORIA) || colHit(COL.SUBCATEGORIA);

      const blocoVals   = sheet.getRange(rowIni, COL.DATA, numRows, (COL.ID_EXTRATO - COL.DATA + 1)).getValues();
      const centrosVals = sheet.getRange(rowIni, COL.CENTRO_CUSTO,   numRows, 1).getValues().flat();
      const tiposVals   = sheet.getRange(rowIni, COL.TIPO,           numRows, 1).getValues().flat();
      const fpVals      = sheet.getRange(rowIni, COL_FP,             numRows, 1).getValues().flat();

      const bufData=[], fmtData=[];
      const bufCat=[];
      const bufParc=[], bufValParc=[];
      const bufCompCaixa=[], fmtCompCaixa=[], notesCompCaixa=[];
      const bufCompCons=[],  fmtCompCons=[];
      const bufStatus=[], bufLiq=[];
      const bufCentro=[], bufTipo=[];
      const bufId=[], bufFP=[];

      const hoje = _today_();
      const idx  = (c) => c - COL.DATA;

      for (let i=0;i<numRows;i++){
        const r   = rowIni + i;
        const row = blocoVals[i].slice();
        const temAlgo = row.some(v=>v!=='' && v!=null);
        if (!temAlgo) continue;

        if (_estaFechadoPorComp_(row)){
          if (![COL.STATUS, COL.LIQUIDACAO, COL.CENTRO_CUSTO, COL.TIPO].some(c => c>=colIni && c<=colFim)){
            _maybeToast_('🔒 Mês fechado: só Status, Liquidação, Centro de Custo e Tipo podem mudar.');
            continue;
          }
          const stNow = String(row[idx(COL.STATUS)]||'').trim();
          const liqNow = row[idx(COL.LIQUIDACAO)];
          if (_isConciliadoStatus_(stNow) && !liqNow) bufLiq.push({r,v:_today_()});
          else if (/pendente/i.test(stNow))          bufLiq.push({r,v:''});
          continue;
        }

        const dtRaw = row[idx(COL.DATA)];
        let d = (dtRaw instanceof Date && !isNaN(dtRaw)) ? _dateOnly_(dtRaw) : parseDateBR_(dtRaw);
        const sub = String(row[idx(COL.SUBCATEGORIA)]||'').trim();
        const det = String(row[idx(COL.DETALHAMENTO)]||'').trim();
        if (!d && (sub||det)) d = _today_();
        if (d){ bufData.push({r,v:d}); fmtData.push({r,fmt:'dd/MM/yyyy'}); row[idx(COL.DATA)]=d; }

        let cat   = String(row[idx(COL.CATEGORIA)]||'').trim();
        let forma = String(row[idx(COL.FORMA)]||'').trim();
        if (sub){
          const catDesejada = mapSubToCat[sub] || '';
          if (cat !== catDesejada){ cat = catDesejada; bufCat.push({r,v:cat}); row[idx(COL.CATEGORIA)]=cat; }
        }else{
          if (colHit(COL.SUBCATEGORIA)){
            sheet.getRange(r, COL.DATA, 1, (COL.ID_EXTRATO - COL.DATA + 1)).clearContent().clearNote().setBackground(null);
            sheet.getRange(r, COL2.COMP_CONSUMO, 1, 1).clearContent();
            sheet.getRange(r, COL_FP,          1, 1).clearContent();
            sheet.getRange(r, COL.DETALHAMENTO).clearDataValidations();
            impactaPrev = true; impactaFaturas = true; continue;
          }
        }

        let parcelas = parseInt(row[idx(COL.PARCELAS)],10);
        if (!isFinite(parcelas) || parcelas<1) parcelas = 1;

        const valorRaw = row[idx(COL.VALOR)];
        let   valorPar = row[idx(COL.VALOR_PARCELADO)];
        const isCartao = ehCartao_(forma);

        const totalCents = toCents_(valorRaw);
        const parcCents  = toCents_(valorPar);
        const valorApagado = (colHit(COL.VALOR) && (valorRaw==='' || valorRaw===null));

        if (valorApagado){
          if (parcCents!==0){
          }else{
            if (parcelas!==1){ parcelas=1; bufParc.push({r,v:1}); }
            if (valorPar!==''){ valorPar=''; bufValParc.push({r,v:''}); }
          }
        }else{
          if (!isCartao){
            if (totalCents !== 0 && parcCents === 0){
              const p0 = fromCents_(_partsCents_(valorRaw, parcelas, null)[0]);
              if (p0!==0){ valorPar=p0; bufValParc.push({r,v:valorPar}); }
            }else if (totalCents===0 && parcCents!==0){
            }else{
              if (valorPar!==''){ valorPar=''; bufValParc.push({r,v:''}); }
              if (parcelas!==1){ parcelas=1; bufParc.push({r,v:1}); }
            }
          }else{
            if (parcelas<1){ parcelas=1; bufParc.push({r,v:1}); }
            if (parcCents === 0){
              const p0 = fromCents_(_partsCents_(valorRaw, parcelas, null)[0]);
              if (p0!==0){ valorPar=p0; bufValParc.push({r,v:valorPar}); }
            }
          }
        }
        row[idx(COL.PARCELAS)]        = parcelas;
        row[idx(COL.VALOR_PARCELADO)] = valorPar;

        if (!d && colHit(COL.DATA)){
          bufCompCaixa.push({ r, v: '' });
          notesCompCaixa.push({ r, note: '' });
          bufCompCons.push({ r, v: '' });
        }
        if (d){
          const compConsumoDate = new Date(d.getFullYear(), d.getMonth(), 1);
          let compCaixaDate = compConsumoDate, compNote = '';
          if (isCartao){
            const cartRow = getCartRow(forma);
            if (cartRow){
              const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]);
              if (venc!=null && ini!=null && fim!=null){
                const {mes,ano} = calcularCicloCartao(d, ini, fim, venc);
                compCaixaDate = new Date(ano, mes-1, 1);
              }else{
                compNote='Ciclo inválido → mês da compra';
              }
            }
          }
          bufCompCaixa.push({r,v:compCaixaDate}); fmtCompCaixa.push({r,fmt:'MM/yyyy'});
          notesCompCaixa.push({ r, note: compNote ? compNote : '' });
          bufCompCons .push({r,v:compConsumoDate}); fmtCompCons .push({r,fmt:'MM/yyyy'});
        }

        let idExt = String(row[idx(COL.ID_EXTRATO)]||'').trim();
        if (!idExt){
          idExt = _gerarIdExtrato_(sheet, r, row, { returnOnly:true });
          bufId.push({ r, v:idExt });
          row[idx(COL.ID_EXTRATO)] = idExt;
        }

        let status = String(row[idx(COL.STATUS)]||'').trim();
        let liq    = row[idx(COL.LIQUIDACAO)];

        if (!isCartao && colHit(COL.LIQUIDACAO)){
          const hasDate = (liq instanceof Date && !isNaN(liq)) || !!parseDateBR_(liq);
          if (hasDate && !_isConciliadoStatus_(status)){ status = "Conciliado"; bufStatus.push({r,v:status}); }
          if (!hasDate &&  _isConciliadoStatus_(status)){ status = "Pendente";   bufStatus.push({r,v:status}); }
        }

        const futuro = d && _dateOnly_(d).getTime() > _dateOnly_(hoje).getTime();
        if (isCartao){
          if (!status){ status="Pendente"; bufStatus.push({r,v:status}); row[idx(COL.STATUS)]=status; }
          if (_isConciliadoStatus_(status) && !liq){ liq=_today_(); bufLiq.push({r,v:liq}); }
        }else{
          const totalAbs = Math.abs(fromCents_(totalCents));
          if (futuro && !_isConciliadoStatus_(status)){
            if (!/pendente/i.test(String(status||""))){ status="Pendente"; bufStatus.push({r,v:status}); row[idx(COL.STATUS)]=status; }
            if (!_isConciliadoStatus_(status)){ liq=''; bufLiq.push({r,v:''}); row[idx(COL.LIQUIDACAO)]=liq; }
          }else{
            if (totalAbs >= EPS_CENT && (!status||status==="")){
              status="Conciliado"; bufStatus.push({r,v:status}); row[idx(COL.STATUS)]=status;
              if (!liq){ liq=_today_(); bufLiq.push({r,v:liq}); }
            }
          }
        }

        if (!String(centrosVals[i]??'').trim()){
          let chosen = centroDefault;
          try{
            if (forma){
              const key = _normLower_(_stripDiacritics_(String(forma)));
              if (cardCentersByName.has(key)) chosen = cardCentersByName.get(key);
            }
          }catch(_){}
          if (chosen) bufCentro.push({r,v:chosen});
        }
        if (!String(tiposVals[i]??'').trim()){
          bufTipo.push({r,v:tiposDefault});
        }

        const fpNew = _rowFingerprint_(row);
        const fpOld = String(fpVals[i]||'');
        if (fpNew !== fpOld){
          bufFP.push({r,v:fpNew});
        }

        if (colHit(COL.DATA, COL.SUBCATEGORIA, COL.DETALHAMENTO, COL.FORMA, COL.PARCELAS, COL.VALOR, COL.VALOR_PARCELADO)){
          try{ _orcamentoAvisarSeUltrapassarLinha_(sheet, r); }catch(_){}
        }
      }

      _setColValuesBatch_(sheet, COL.DATA,            bufData);
      _setColFormatBatch_( sheet, COL.DATA,           fmtData);
      _setColValuesBatch_(sheet, COL.CATEGORIA,       bufCat);
      _setColValuesBatch_(sheet, COL.PARCELAS,        bufParc);
      _setColValuesBatch_(sheet, COL.VALOR_PARCELADO, bufValParc);
      _setColValuesBatch_(sheet, COL.COMPETENCIA,     bufCompCaixa);
      _setColFormatBatch_( sheet, COL.COMPETENCIA,    fmtCompCaixa);
      _setColNotesBatch_(  sheet, COL.COMPETENCIA,    notesCompCaixa);
      _setColValuesBatch_(sheet, COL2.COMP_CONSUMO,   bufCompCons);
      _setColFormatBatch_( sheet, COL2.COMP_CONSUMO,  fmtCompCons);
      _setColValuesBatch_(sheet, COL.STATUS,          bufStatus);
      _setColValuesBatch_(sheet, COL.LIQUIDACAO,      bufLiq);
      _setColValuesBatch_(sheet, COL.CENTRO_CUSTO,    bufCentro);
      _setColValuesBatch_(sheet, COL.TIPO,            bufTipo);
      _setColValuesBatch_(sheet, COL.ID_EXTRATO,      bufId);
      _setColValuesBatch_(sheet, COL_FP,              bufFP);

      if (precisaAtualizarMenusFixos){
        atualizarMenuCategoriasLancamento(sheet, categoriasLista);
        atualizarMenuSubcategoriasLancamento(sheet, subsCfg);
      }
      try{ atualizarMenusDinamicos(sheet, shCfg, rowIni, numRows); }catch(_){}

      doRecalcsNow_({ faturas: impactaFaturas, previsao: impactaPrev, resumo: true });
      return;
    }

    /* ========== ABAS MENSAIS (conveniências) ========== */
    if (MESES.map(_normLower_).includes(nomeAbaNorm)){
      if (colIni===3 && rowIni>=20 && rowIni<=50){
        const editedValue = range.getValue();
        if (!editedValue) sheet.getRange(rowIni,2,1,5).clearContent();
        else if (!sheet.getRange(rowIni,2).getValue()) sheet.getRange(rowIni,2).setValue(_today_());
      }
      if (colIni===2){
        try{
          const start=20, linhas=31;
          const datas  = sheet.getRange(start,2,linhas,1).getValues();
          const valores= sheet.getRange(start,3,linhas,1).getValues();
          let mudou=false;
          for(let i=0;i<linhas;i++){
            if(!datas[i][0] && valores[i][0]){ datas[i][0]=_today_(); mudou=true; }
          }
          if(mudou) sheet.getRange(start,2,linhas,1).setValues(datas);
        }catch(_){}
      }
      return;
    }

  }catch(err){
    Logger.log('onEdit error: ' + (err && err.message ? err.message : err));
  }
}


/***************************************************************
 * FINANCEIRO FAMILIAR — SCRIPT UNIFICADO (PARTE 2 — 2A + 2B)
 * HOTFIX: 28/10/2025 (idempotente, single-user friendly)
 *
 * 2A: Faturas & Resumo (infra), limites/exposição, geração de faturas,
 *     resumo (pagamento parcial) — com cabeçalhos protegidos.
 * 2B: Previsão/Orçamento, Projeções (mediana/winsor), Auditorias.
 *
 * Depende da PARTE 1 (helpers, ABAS/COL/COL2/COL_FP, etc.).
 ***************************************************************/

/* ===================== Helpers locais do bloco (com hotfixes) ===================== */

// Cria a aba só se NÃO existir (idempotente e "race-safe")
function _safeEnsureSheet_(name, alternates){
  const ss = SS_();
  let sh = getSheetSmart_(name, alternates||[]);
  if (sh) return sh;
  try {
    return ss.insertSheet(name);
  } catch(e){
    const msg = String(e && e.message || '');
    if (msg.includes('Já existe uma página') || /already exists/i.test(msg)){
      sh = getSheetSmart_(name, alternates||[]);
      if (sh) return sh;
    }
    throw e;
  }
}

// Retorna valorParcRaw se presente; senão usa fallback (já normalizado)
function _parcelaOuFallback_(valorParcRaw, fallback){
  if (valorParcRaw != null && String(valorParcRaw).trim() !== ''){
    try { return fromCents_(toCents_(valorParcRaw)); } catch(_){}
  }
  return fallback;
}

// Proteção de cabeçalho (reforçada com editor efetivo)
function _protectHeaderRow_(sh, opt){
  if (features_().NO_PROTECT) return;
  try{
    const DESC = 'Cabeçalhos protegidos';
    const dp   = PropertiesService.getDocumentProperties();
    const key  = 'HDR_PROT_SIG_' + sh.getSheetId();
    const want = JSON.stringify({ desc:DESC, cols: sh.getMaxColumns(), v:2, row:1 });

    const prots = (sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)||[])
      .filter(p => (p.getDescription()||'') === DESC);

    const ok = prots.some(p => {
      try{
        const rg = p.getRange();
        return rg && rg.getRow()===1 && rg.getNumRows()===1 &&
               rg.getColumn()===1 && rg.getNumColumns()===sh.getMaxColumns();
      }catch(_){ return false; }
    });

    const force = opt && opt.force === true;
    if (ok && !force && dp.getProperty(key) === want) return;

    prots.forEach(p => { try{ p.remove(); }catch(_){ } });
    const rng  = sh.getRange(1,1,1, Math.max(1, sh.getMaxColumns()));
    const prot = rng.protect().setDescription(DESC);
    try { prot.setWarningOnly(false); } catch(_){}
    try { prot.removeEditors(prot.getEditors()); } catch(_){}
    try { prot.addEditor(Session.getEffectiveUser()); } catch(_){}
    dp.setProperty(key, want);
  }catch(_){}
}

/* ===================== [2A] FATURAS & RESUMO — infra ===================== */

function _ensureResumoHeaders_(){
  const sh = _safeEnsureSheet_(ABAS.RESUMO_FATURAS, ['Resumo de Faturas']);

  const hdr=[
    "Cartão","Mês","Total (R$)","Pendentes (R$)","Conciliadas (R$)",
    "Último Status","CHAVE","Utilização (%)","Exposição Futura (R$)"
  ];

  // Cabeçalho – só reescreve se diferente
  const cur=sh.getRange(1,1,1,hdr.length).getValues()[0].map(v=>String(v||"").trim());
  let changedHeader=false;
  for(let i=0;i<hdr.length;i++){ if((cur[i]||"")!==hdr[i]){ changedHeader=true; break; } }
  if(changedHeader){ sh.getRange(1,1,1,hdr.length).setValues([hdr]); }

  // Formatação idempotente por “assinatura” (leve)
  try{
    const dp = PropertiesService.getDocumentProperties();
    const key= 'HDR_FMT_RESUMO_' + sh.getSheetId();
    const sig= JSON.stringify({ nfB:'MM/yyyy', nfCE:'R$ #,##0.00', nfH:'0.00%', nfI:'R$ #,##0.00', hideC7:true, freeze:[1,1], v:2 });
    if (dp.getProperty(key) !== sig || changedHeader){
      sh.getRange("B:B").setNumberFormat("MM/yyyy");
      sh.getRange("C:E").setNumberFormat("R$ #,##0.00");
      sh.getRange("H:H").setNumberFormat('0.00%');
      sh.getRange("I:I").setNumberFormat("R$ #,##0.00");
      try{ sh.hideColumns(7); }catch(_){}
      try{ sh.setFrozenRows(1); sh.setFrozenColumns(1); }catch(_){}
      dp.setProperty(key, sig);
    }
  }catch(_){}

  _protectHeaderRow_(sh);
  return sh;
}

function _ensureFaturasHeaders_(){
  const sh = _safeEnsureSheet_(ABAS.FATURAS_CARTAO, ['Faturas de Cartao']);

  const hdr=["Cartão","Mês Referência","Data Vencimento","Valor Fatura","Status Pagamento","Data Pagamento","CHAVE","Valor Pago","Encargos","Créditos","Observações"];

  // Cabeçalho — só reescreve se diferente
  const cur=sh.getRange(1,1,1,hdr.length).getValues()[0].map(v=>String(v||"").trim());
  let diff=false; for(let i=0;i<hdr.length;i++){ if(hdr[i]!==cur[i]){ diff=true; break; } }
  if(diff){ sh.getRange(1,1,1,hdr.length).setValues([hdr]); }

  // Formatos + DV SEM depender de assinatura (evita DV “congelada”)
  try{
    sh.getRange("B:B").setNumberFormat("MM/yyyy");
    sh.getRange("C:C").setNumberFormat("dd/MM/yyyy");
    sh.getRange("D:D").setNumberFormat("R$ #,##0.00");
    sh.getRange("F:F").setNumberFormat("dd/MM/yyyy");
    sh.getRange("H:J").setNumberFormat("R$ #,##0.00");
    try{ sh.hideColumns(7); }catch(_){}
    try{ sh.setFrozenRows(1); sh.setFrozenColumns(1); }catch(_){}
  }catch(_){}

  // DV Cartão (A2:A) — sempre reavalia a partir da Config
  try{
    const cfg=_getCfg_();
    if(cfg){
      const rows=_rangeCartoes_(cfg);
      if(rows.length){
        const map=_cfgMap_(cfg);
        const cartoesRng=cfg.getRange(map.CARTOES_FIRST_ROW, map.CARTOES_FIRST_COL, rows.length, 1);
        _applyDVRangeIfChanged_(sh.getRange("A2:A"), cartoesRng, false);
      } else {
        sh.getRange("A2:A").clearDataValidations();
      }
    }
  }catch(_){}

  // DV Status (E2:E)
  try{
    const statusList=["Pendente","Conciliado","Cancelado"];
    _applyDVIfChanged_(sh.getRange("E2:E"), statusList, false);
  }catch(_){}

  _protectHeaderRow_(sh);
  return sh;
}

/* ===================== Limites & Exposição ===================== */

function _cardKeyFromName_(nome){
  const s=String(nome||"").toLowerCase().trim();
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,s);
  const hex=bytes.map(b=>('0'+(b&0xFF).toString(16)).slice(-2)).join('').slice(0,6).toUpperCase();
  return `CARD#${hex}`;
}

// opcional (use somente se adotar também na geração/CHAVE)
function _cardKeyFromNameAndCycle_(nome, ini, fim, venc){
  const s = `${String(nome||'').trim().toLowerCase()}|${ini}|${fim}|${venc}`;
  const h = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8)
    .map(b => (b+256).toString(16).slice(-2)).join('');
  return 'CARD#' + h.slice(0,10);
}


function _limitePorCartao_(cfg){
  const rows = _rangeCartoes_(cfg);
  const mapa = new Map();
  rows.forEach(r=>{
    const nome = String(r[0]||"").trim();
    let lim = 0;
    const raw = r[4];
    if (raw != null) {
      if (typeof raw === 'number') lim = raw;             // já numérico
      else try { lim = fromCents_(toCents_(raw)); } catch(_) { lim = 0; } // parse "R$ 1.550,00"
    }
    if (nome) mapa.set(nome, lim);
  });
  return mapa;
}

function _exposicaoFuturaPorCartao_(shLanc, cfg, includeCurrent /*=false*/){
  includeCurrent = Boolean(includeCurrent);
  const last=shLanc.getLastRow(); const mapa=new Map(); if(last<4) return mapa;
  const rows=shLanc.getRange(4, COL.DATA, last-3, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues();
  const hoje=_today_(); const baseYear=hoje.getFullYear(), baseMonth=hoje.getMonth()+1;

  for(let i=0;i<rows.length;i++){
    const [data, , , , forma, parcelasRaw, valorRaw, valorParcRaw]=rows[i];
    const dt=(data instanceof Date && !isNaN(data))?data:parseDateBR_(data); if(!dt) continue;

    const cartRow=_findCartaoRow_(forma); if(!cartRow) continue;
    const cartaoNome=String(cartRow[0]||"").trim();
    const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]); if(venc==null||ini==null||fim==null) continue;

    let n=parseInt(parcelasRaw,10); if(!isFinite(n)||n<=0) n=1;
    const total=Number(valorRaw)||0; const parts=_parcelasExatas_(total, n);

    for(let p=0;p<n;p++){
      const dataParc=new Date(dt); dataParc.setMonth(dataParc.getMonth()+p);
      const { mes:mFat, ano:aFat }=calcularCicloCartao(dataParc, ini, fim, venc);

      const isFuture = (aFat>baseYear) || (aFat===baseYear && mFat>baseMonth);
      const isCurrent= (aFat===baseYear && mFat===baseMonth);
      if(!(isFuture || (includeCurrent && isCurrent))) continue;

      const vParc = _parcelaOuFallback_(valorParcRaw, parts[p]);
      if (_eqNum_(vParc, 0)) continue;
      mapa.set(cartaoNome, (mapa.get(cartaoNome)||0)+vParc);
    }
  }
  // Arredondamento consistente na fonte
  for (const [k,v] of mapa.entries()) mapa.set(k, _r2(v));
  return mapa;
}

function _exposicaoFuturaBuckets_(shLanc){
  const last=shLanc.getLastRow();
  const out={ d30:0, d90:0, d180:0, dInf:0 }; if(last<4) return out;
  const rows=shLanc.getRange(4, COL.DATA, last-3, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues();
  const hoje=_today_();

  for(let i=0;i<rows.length;i++){
    const [data, , , , forma, parcelasRaw, valorRaw, valorParcRaw]=rows[i];
    const dt=(data instanceof Date && !isNaN(data))?data:parseDateBR_(data); if(!dt) continue;

    const cartRow=_findCartaoRow_(forma); if(!cartRow) continue;
    const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]);
    if(venc==null||ini==null||fim==null) continue;

    let n=parseInt(parcelasRaw,10); if(!isFinite(n)||n<=0) n=1;
    const total=Number(valorRaw)||0; const parts=_parcelasExatas_(total, n);

    for(let p=0;p<n;p++){
      const dParc=new Date(dt); dParc.setMonth(dParc.getMonth()+p);
      const info=calcularCicloCartao(dParc, ini, fim, venc);
      const vencDate=info && (info.venc || info.vencDate);
      if(!(vencDate instanceof Date) || isNaN(vencDate)) continue; // guarda: sem NaN

      const v = _parcelaOuFallback_(valorParcRaw, parts[p]);
      if (_eqNum_(v, 0)) continue;
      const dias=Math.ceil((_dateOnly_(vencDate)-_dateOnly_(hoje))/(24*3600*1000));
      if(dias<0) continue;
      if(dias<=30) out.d30+=v;
      else if(dias<=90) out.d90+=v;
      else if(dias<=180) out.d180+=v;
      else out.dInf+=v;
    }
  }
  Object.keys(out).forEach(k=>out[k]=_r2(out[k]));
  return out;
}

/* ===================== FATURAS (geração) ===================== */

/* ==== compat: acha linha do cartão mesmo com variação de nome na Forma ==== */
function _findCartaoRowCompat_(forma, cfg, cachedRows){
  // 1) tenta o helper oficial (se existir e casar)
  try {
    if (typeof _findCartaoRow_ === 'function') {
      const r = _findCartaoRow_(forma);
      if (r) return r;
    }
  } catch(_){}

  // 2) fallback: normaliza e tenta match aproximado nos cartões da Config
  const rows = cachedRows && cachedRows.length ? cachedRows : _rangeCartoes_(cfg);
  const norm = s => _normLower_(_stripDiacritics_(String(s||'').trim()));
  const target = norm(forma);

  // aceita igual, contains e endsWith (permite "Cartão X • Visa", "Visa - X", etc.)
  for (const r of rows){
    const nome = String(r[0]||'').trim();
    if (!nome) continue;
    const n = norm(nome);
    if (n === target || n.includes(target) || target.includes(n) || target.endsWith(n) || n.endsWith(target)){
      return r;
    }
  }
  return null;
}

/* ===================== FATURAS (geração) — COM FALLBACK ===================== */
function gerarFaturasDeCartao_(useLock=true){
  let lock;
  if(useLock){
    lock=LockService.getDocumentLock();
    let ok=lock.tryLock(5000);
    for(let i=0; !ok && i<2; i++){ Utilities.sleep(400); ok=lock.tryLock(5000); }
    if(!ok){ _maybeToast_("⏳ Outro processo em andamento..."); return; }
  }
  try{
    const tz=_tz_();
    const shLanc=_getLanc_();
    const shCfg =_getCfg_();
    const shFat=_ensureFaturasHeaders_();
    if(!shLanc||!shCfg||!shFat) return;

    const cartoes=_rangeCartoes_(shCfg);
    if(!cartoes.length){ _maybeToast_("⚠️ Nenhum cartão em Configurações!"); return; }

    const ultima=shLanc.getLastRow();
    if(ultima<4){ _maybeToast_("ℹ️ Sem lançamentos para gerar faturas."); return; }

    const rows=shLanc.getRange(4, COL.DATA, ultima-3, COL.STATUS - COL.DATA + 1).getValues(); // B..J
    const mapa=new Map();
    const errosCiclo=new Set();

    function addToMap(cardKey, cartaoNome, mesRefDate, dtVenc, valor){
      const chave=`${cardKey}||${Utilities.formatDate(mesRefDate, tz, "MM/yyyy")}`;
      if(!mapa.has(chave)) mapa.set(chave, {cartaoNome, mesRef:mesRefDate, dtVenc, total:0});
      const o=mapa.get(chave); o.total+=(Number(valor)||0);
    }

    for(let i=0;i<rows.length;i++){
      const [data, , , , forma, parcelasRaw, valorRaw, valorParcRaw]=rows[i];
      if(!data||!forma) continue;

      // 🔧 aqui entra o fallback robusto:
      let cartRow = _findCartaoRowCompat_(forma, shCfg, cartoes);
      if(!cartRow) continue;

      const cartaoNome=String(cartRow[0]||"").trim();
      const cardKey=_cardKeyFromName_(cartaoNome);

      const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]);
      if(venc==null||ini==null||fim==null){ errosCiclo.add(cartaoNome); continue; }

      const dt=(data instanceof Date&&!isNaN(data))?data:parseDateBR_(data);
      if(!dt) continue;

      let parcelas=parseInt(parcelasRaw,10); if(!isFinite(parcelas)||parcelas<=0) parcelas=1;
      const total=Number(valorRaw)||0;
      const parts=_parcelasExatas_(total, parcelas);

      for(let p=0;p<parcelas;p++){
        const dataParc=new Date(dt); dataParc.setMonth(dataParc.getMonth()+p);
        const { mes:mFat, ano:aFat, venc: vencDate }=calcularCicloCartao(dataParc, ini, fim, venc);
        const mesRefDate=new Date(aFat, mFat-1, 1);
        const vParc = _parcelaOuFallback_(valorParcRaw, parts[p]);
        if (_eqNum_(vParc, 0)) continue;
        addToMap(cardKey, cartaoNome, mesRefDate, vencDate, vParc);
      }
    }

    // upsert na aba Faturas
    const lastFat=shFat.getLastRow();
    const existentes=(lastFat>=2)? shFat.getRange(2,7,lastFat-1,1).getValues().flat() : [];
    const idx=new Map(); existentes.forEach((k,i)=>{ if(k) idx.set(k, i+2); });

    let atualizados=0; const toInsert=[];
    for(const [chave, o] of mapa.entries()){
      const r=idx.get(chave);
      const linha=[o.cartaoNome, o.mesRef, o.dtVenc, _r2(o.total), "", "", chave, "", "", "", ""]; // A..K
      if(r){
        const cur=shFat.getRange(r,1,1,linha.length).getValues()[0];
        linha[4]=cur[4]||linha[4]; // Status
        linha[5]=cur[5]||linha[5]; // Data Pagamento
        linha[7]=cur[7]||linha[7]; // Pago
        linha[8]=cur[8]||linha[8]; // Encargos
        linha[9]=cur[9]||linha[9]; // Créditos
        linha[10]=cur[10]||linha[10]; // Obs

        const igual= _sameDate_(cur[1], linha[1]) && _sameDate_(cur[2], linha[2]) &&
                     _eqNum_(cur[3], linha[3]) && String(cur[6]||"")==linha[6] &&
                     String(cur[0]||"")==linha[0];
        if(!igual){
          shFat.getRange(r,1,1,linha.length).setValues([linha]);
          try{
            shFat.getRange(r,2).setNumberFormat("MM/yyyy");
            shFat.getRange(r,3).setNumberFormat("dd/MM/yyyy");
            shFat.getRange(r,4).setNumberFormat("R$ #,##0.00");
          }catch(_){ }
          atualizados++;
        }
      } else {
        toInsert.push(linha);
      }
    }

    if(toInsert.length){
      const start=shFat.getLastRow()+1;
      shFat.getRange(start, 1, toInsert.length, 11).setValues(toInsert);
      try{
        shFat.getRange(start,2,toInsert.length,1).setNumberFormat("MM/yyyy");
        shFat.getRange(start,3,toInsert.length,1).setNumberFormat("dd/MM/yyyy");
        shFat.getRange(start,4,toInsert.length,1).setNumberFormat("R$ #,##0.00");
      }catch(_){ }
    }

    // limpar obsoletos e ordenar
    try{
      const keysAtuais=new Set(mapa.keys());
      const aExcluir=[];
      idx.forEach((row, key)=>{ if(!keysAtuais.has(key)) aExcluir.push(row); });
      aExcluir.sort((a,b)=>b-a).forEach(r=>shFat.deleteRow(r));
    }catch(_){}
    try{
      const n=shFat.getLastRow();
      if(n>1) shFat.getRange(2,1,n-1,11).sort([
        {column:1, ascending:true}, // Cartão
        {column:2, ascending:true}, // Mês
        {column:3, ascending:true}  // Venc
      ]);
    }catch(_){}

    if(errosCiclo.size){
      const nomes=Array.from(errosCiclo), max=8;
      const msg=nomes.length>max?`${nomes.slice(0,max).join(', ')} +${nomes.length-max}`:nomes.join(', ');
      _maybeToast_(`⚠️ Ciclo inválido: ${msg} (confira venc/início/fim).`);
    }

    try{ _protectHeaderRow_(_ensureFaturasHeaders_()); }catch(_){}

    _maybeToast_(`✅ Faturas atualizadas. +${toInsert.length} inseridas / ${atualizados} atualizadas`);
  }catch(e){
    _maybeToast_("⚠️ Erro em gerarFaturasDeCartao_: "+(e&&e.message?e.message:e));
  } finally {
    try{ if(useLock && lock) lock.releaseLock(); }catch(_){}
  }
}

/* ===================== RESUMO (pagamento parcial) ===================== */

// Normaliza Coluna B (Mês) para Date antes de ordenar
function normalizeResumoMesColB_(shRes){
  const last=shRes.getLastRow(); if(last<2) return;
  const rg  = shRes.getRange(2,2,last-1,1);
  const vals= rg.getValues();
  for(let i=0;i<vals.length;i++){
    const c=vals[i][0];
    if(c instanceof Date && !isNaN(c)) continue;
    const s=String(c||'').trim();
    if(/^\d{2}\/\d{4}$/.test(s)){
      const [mm,yyyy]=s.split('/').map(Number);
      vals[i][0]=new Date(yyyy, mm-1, 1);
    } else if (s){
      const p=parseDateBR_(s);
      if(p) vals[i][0]=new Date(p.getFullYear(), p.getMonth(), 1);
    }
  }
  rg.setValues(vals);
  try{ rg.setNumberFormat('MM/yyyy'); }catch(_){}
}

// Atualizar Resumo de Faturas — 

function atualizarResumoFaturas_(useLock=true){
  let lock;
  if(useLock){
    lock=LockService.getDocumentLock();
    let ok=lock.tryLock(5000);
    for(let i=0; !ok && i<2; i++){ Utilities.sleep(400); ok=lock.tryLock(5000); }
    if(!ok){ _maybeToast_("⏳ Outro processo em andamento..."); return; }
  }
  try{
    const tz=_tz_();
    const shFat=_ensureFaturasHeaders_(), shRes=_ensureResumoHeaders_();
    if(!shFat||!shRes){ _maybeToast_("❌ Abas não encontradas."); return; }

    const last=shFat.getLastRow();
    const dados=(last>=2)? shFat.getRange(2,1,last-1,11).getValues() : [];

    // agrega por CHAVE (cartão+mês)
    const mapa=new Map();
    for(const r of dados){
      const cartao=String(r[0]||"").trim();
      const mesCell=r[1];
      const mesDate=(mesCell instanceof Date&&!isNaN(mesCell))? new Date(mesCell.getFullYear(), mesCell.getMonth(),1):null;
      const mesStr=mesDate? Utilities.formatDate(mesDate, tz, "MM/yyyy") : String(mesCell||"").trim();
      const valFatura=Number(r[3]||0);
      const status=String(r[4]||"").trim();
      const dtPag=r[5];
      const chave=String(r[6]||"").trim();
      const valPago=Number(r[7]||0);
      const encargos=Number(r[8]||0);
      const creditos=Number(r[9]||0);

      if(!(cartao && mesStr)) continue;
      const k=chave || `${_cardKeyFromName_(cartao)}||${mesStr}`;
      if(!mapa.has(k)) mapa.set(k, { cartao, mesStr, mesDate, total:0, pagos:0, encargos:0, creditos:0, lastStatus:"", lastPaidAt:null });

      const o=mapa.get(k);
      o.total    += valFatura;
      o.pagos    += Math.max(valPago,0);
      o.encargos += Math.max(encargos,0);
      o.creditos += Math.max(creditos,0);

      if(status){
        const paidAt = (dtPag instanceof Date && !isNaN(dtPag)) ? dtPag : null;
        if(!o.lastPaidAt || (paidAt && paidAt > o.lastPaidAt)){
          o.lastStatus = status;
          o.lastPaidAt = paidAt;
        }
        if (_isConciliadoStatus_(status) && !paidAt && !o.lastPaidAt){
          o.lastStatus = status;
        }
      }
    }

    // limites/exposição
    const cfg=_getCfg_();
    const limites=_limitePorCartao_(cfg);
    const shLanc=_getLanc_();
    const expoFutura=_exposicaoFuturaPorCartao_(shLanc, cfg /* includeCurrent=false */);

    // Comparador com tolerância (~0,05 p.p.)
    const utilEq = (a,b) => {
      const na = (a===''||a==null) ? null : Number(a);
      const nb = (b===''||b==null) ? null : Number(b);
      if (na==null && nb==null) return true;
      if (na==null || nb==null) return false;
      return Math.abs(na - nb) <= 0.0005;
    };

    const lastRes=shRes.getLastRow();
    const existentes=(lastRes>=2)? shRes.getRange(2,7,lastRes-1,1).getValues().flat() : [];
    const idxPorChave=new Map(); existentes.forEach((ch,i)=>{ if(ch) idxPorChave.set(ch, i+2); });

    const toInsert=[];
    for(const [key, o] of mapa.entries()){
      const limite=Number(limites.get(o.cartao)||0);
      const liquido = _r2( (o.total + o.encargos - o.creditos) );
      const pendente= _r2( Math.max(liquido - o.pagos, 0) );
      const conciliado=_r2( Math.min(o.pagos, liquido) );
      const utilizacao=(limite>0)? (liquido/limite) : null;
      const expo=_r2(Number(expoFutura.get(o.cartao)||0));

      const mesDateFinal=o.mesDate || (function(){ const m=o.mesStr.split('/'); if(m.length===2){ const mm=+m[0], yyyy=+m[1]; return new Date(yyyy, mm-1, 1); } return null; })();
      const linha=[ o.cartao, mesDateFinal, liquido, pendente, conciliado, o.lastStatus||"", key, utilizacao, expo ];
      const row=idxPorChave.get(key);
      if(row){
        const atual=shRes.getRange(row,1,1,9).getValues()[0];
        const statusAtual=String(atual[5]||"");
        const finalStatus=_isConciliadoStatus_(statusAtual)?statusAtual:(o.lastStatus||"");
        const igual=(String(atual[0]||"")==linha[0]) && _sameDate_(atual[1],linha[1]) &&
          _eqNum_(atual[2],linha[2]) && _eqNum_(atual[3],linha[3]) && _eqNum_(atual[4],linha[4]) &&
          (String(statusAtual)==finalStatus) && (String(atual[6]||"")==linha[6]) &&
          utilEq(atual[7], linha[7]) && _eqNum_(atual[8],linha[8]);
        if(!igual){
          linha[5]=finalStatus;
          shRes.getRange(row,1,1,9).setValues([linha]);
          try{
            shRes.getRange(row,2).setNumberFormat("MM/yyyy");
            shRes.getRange(row,3).setNumberFormat("R$ #,##0.00");
            shRes.getRange(row,4).setNumberFormat("R$ #,##0.00");
            shRes.getRange(row,5).setNumberFormat("R$ #,##0.00");
            shRes.getRange(row,8).setNumberFormat("0.00%");
            shRes.getRange(row,9).setNumberFormat("R$ #,##0.00");
          }catch(_){ }
        }
      } else toInsert.push(linha);
    }
    if(toInsert.length){
      const start=shRes.getLastRow()+1;
      shRes.getRange(start, 1, toInsert.length, 9).setValues(toInsert);
      try{
        shRes.getRange(start,2,toInsert.length,1).setNumberFormat("MM/yyyy");
        shRes.getRange(start,3,toInsert.length,1).setNumberFormat("R$ #,##0.00");
        shRes.getRange(start,4,toInsert.length,1).setNumberFormat("R$ #,##0.00");
        shRes.getRange(start,5,toInsert.length,1).setNumberFormat("R$ #,##0.00");
        shRes.getRange(start,8,toInsert.length,1).setNumberFormat("0.00%");
        shRes.getRange(start,9,toInsert.length,1).setNumberFormat("R$ #,##0.00");
      }catch(_){ }
    }

    // Normaliza Mês (B) para Date e só então ordena
    normalizeResumoMesColB_(shRes);

    // Remoção de obsoletos + Ordenação
    try{
      const keysAtuais=new Set(mapa.keys()); const aExcluir=[];
      for(const [k,row] of idxPorChave.entries()){ if(!keysAtuais.has(k)) aExcluir.push(row); }
      aExcluir.sort((a,b)=>b-a).forEach(r=>shRes.deleteRow(r));
    }catch(_){ }
    try{
      const n=shRes.getLastRow(); if(n>1) shRes.getRange(2,1,n-1,9)
        .sort([{column:1, ascending:true}, {column:2, ascending:true}]);
    }catch(_){ }

    try{ _protectHeaderRow_(_ensureResumoHeaders_()); }catch(_){}
    _maybeToast_(`✅ Resumo de Faturas atualizado (${mapa.size} grupo(s)).`);
    try{ ensureCondFormatResumoUtil_(); }catch(_){}
    try{ alertasUtilizacaoEExposicao_(); }catch(_){}
  }catch(err){ _maybeToast_("⚠️ Erro: "+(err&&err.message?err.message:err)); }
  finally{ try{ if(useLock && lock) lock.releaseLock(); }catch(_){ } }
}

/***************************************************************
 * [2B] PREVISÃO / ORÇAMENTO / PROJEÇÕES / AUDITORIAS
 ***************************************************************/

const MAX_ITENS_PREV = 211;

function _dedupCaseAccent_(arr){
  const seen=new Set(); const out=[];
  for(const s of (arr||[])){
    const k=_normLower_(s);
    if(k && !seen.has(k)){ seen.add(k); out.push(s); }
  }
  return out;
}

function preencherPrevisaoDeGastos(){
  const cfg=_getCfg_(), prev=getSheetSmart_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']);
  if(!cfg||!prev) return;

  const last=_cfgLastRow_(cfg);
  const detalhamentos = cfg.getRange(3,5,last-2,1).getValues().flat().filter(String);
  const subcategorias  = cfg.getRange(3,2,last-2,1).getValues().flat().filter(String);
  const subcatsComDet  = cfg.getRange(3,6,last-2,1).getValues().flat().filter(String);
  const subcatsSemDet  = subcategorias.filter(s => !detalhamentos.includes(s) && !subcatsComDet.includes(s));

  // FIX: usar spread correto
  const listaFinal = _dedupCaseAccent_([...detalhamentos, ...subcatsSemDet]).slice(0, MAX_ITENS_PREV);

  const start=2, max=MAX_ITENS_PREV, n=listaFinal.length;
  if (n) prev.getRange(start,1,n,1).setValues(listaFinal.map(x=>[x]));
  if (n<max) prev.getRange(start+n,1,max-n,1).clearContent();
}

function preencherColunaB(){
  const prev = getSheetSmart_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']);
  const lanc = _getLanc_();
  const cfg  = _getCfg_();
  if (!prev || !lanc || !cfg) return;

  const anoRef   = getAnoRef_(cfg) || (new Date()).getFullYear();
  const compMode = getCompMode_(cfg); // 'CAIXA' | 'CONSUMO'

  // Itens (A2:A)
  const lastPrev = Math.max(prev.getLastRow(), 2);
  const itens    = prev.getRange(2,1,lastPrev-1,1).getValues().flat().map(s => String(s||'').trim());
  if (!itens.length){
    try { prev.getRange('B2:B').clearContent(); } catch(_) {}
    return;
  }

  // --- Normalização unificada de Detalhamentos oficiais ---
  const lastCfg = _cfgLastRow_(cfg);
  const detOfRaw = (lastCfg>2 ? cfg.getRange(3,5,lastCfg-2,1).getValues() : []).flat().filter(Boolean);
  const detsOficiaisNorm = new Set(detOfRaw.map(x => _normLower_(_stripDiacritics_(String(x)))));
  const norm = s => _normLower_(_stripDiacritics_(String(s||'')));

  // Soma por item em CENTAVOS
  const somaCents = new Map();
  const addC = (item, cents) => {
    const k = String(item||'').trim();
    if (!k || !(Number(cents)>0)) return;
    somaCents.set(k, (somaCents.get(k)||0) + Number(cents));
  };

  const lastLanc = lanc.getLastRow();
  if (lastLanc >= 4){
    const rows = lanc.getRange(4, COL.DATA, lastLanc-3, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues();
    for (const r of rows){
      const [dataRaw, sub, det, , forma, parcelasRaw, valorRaw, valorParcRaw] = r;

      const detStr = String(det||'').trim();
      const subStr = String(sub||'').trim();
      const item   = detsOficiaisNorm.has(norm(detStr)) ? detStr : subStr;
      if (!item) continue;

      const dt = (dataRaw instanceof Date && !isNaN(dataRaw)) ? dataRaw : parseDateBR_(dataRaw);
      if (!dt) continue;

      // valores em centavos
      const totalCents = toCents_(valorRaw);
      const parcCents  = toCents_(valorParcRaw);
      let n = parseInt(parcelasRaw, 10); if (!isFinite(n) || n <= 0) n = 1;

      if (compMode === 'CONSUMO'){
        if (dt.getFullYear() === anoRef && totalCents > 0) addC(item, totalCents);
        continue;
      }

      // CAIXA
      const cartRow = _findCartaoRow_(forma);
      if (!cartRow){
        // --- Modelo B não-cartão: VALOR=0 e VALOR_PARCELADO>0 → soma n*parcela no próprio ano ---
        if (dt.getFullYear() === anoRef){
          if (totalCents === 0 && parcCents !== 0){
            addC(item, parcCents * n);
          } else if (totalCents > 0){
            addC(item, totalCents);
          }
        }
        continue;
      }

      // Cartão
      const venc = _sanitizaDia_(cartRow[1]),
            ini  = _sanitizaDia_(cartRow[2]),
            fim  = _sanitizaDia_(cartRow[3]);

      if (venc==null || ini==null || fim==null){
        if (dt.getFullYear() === anoRef && totalCents > 0) addC(item, totalCents);
        continue;
      }

      // monta parcelas exatas em centavos
      const base  = Math.floor((totalCents>0?totalCents:0) / n);
      const resto = (totalCents>0?totalCents:0) - base*n;
      const partsCents = Array.from({length:n}, (_,i)=> base + (i<resto ? 1 : 0));

      for (let p=0; p<n; p++){
        const dParc = new Date(dt); dParc.setMonth(dParc.getMonth() + p);
        const { mes, ano } = calcularCicloCartao(dParc, ini, fim, venc);
        if (ano !== anoRef) continue;

        // prioriza valorParcRaw quando existir
        let vParcCents = (parcCents!==0 && parcCents!=null) ? parcCents : partsCents[p];
        if (vParcCents > 0) addC(item, vParcCents);
      }
    }
  }

  // Escreve B2:B
  const out = itens.map(it => [ _r2(fromCents_(somaCents.get(String(it||'').trim()) || 0)) ]);
  const destino = prev.getRange(2,2,out.length,1);
  destino.setValues(out);
  try { destino.setNumberFormat('R$ #,##0.00'); } catch(_) {}
}

/* ========= Estrutura da aba PREVISÃO (garante criação + formatos) ========= */
function ensureOrcamentoEstrutura_(){
  const sh = _safeEnsureSheet_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']);
  if (!sh) return;

  const desired = [
    'Categoria',            // A
    'Acum. Ano (R$)',       // B
    'Meta Mês (R$)',        // C
    'Saldo Anual (R$)',     // D
    'Desvio Anual (%)',     // E
    'Desvio Mês (%)',       // F
    'Prioridade',           // G
    'Barra Mês',            // H
    'Restante Mês (%)',     // I
    'Consumido Mês (%)',    // J
    'Status Mês'            // K
  ];

  const hdrRange = sh.getRange(1,1,1,desired.length);
  const cur = hdrRange.getValues()[0].map(v=>String(v||'').trim());
  let needsRename = false;
  for (let i=0;i<desired.length;i++){ if(cur[i] !== desired[i]){ needsRename = true; break; } }
  if (needsRename){ sh.getRange(1,1,1,desired.length).setValues([desired]); }

  try {
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('R$ #,##0.00');
    sh.getRange('C:C').setNumberFormat('R$ #,##0.00');
    sh.getRange('D:D').setNumberFormat('R$ #,##0.00');
    sh.getRange('E:E').setNumberFormat('0.00%');
    sh.getRange('F:F').setNumberFormat('0.00%');
    sh.getRange('I:I').setNumberFormat('0.00%');
    sh.getRange('J:J').setNumberFormat('0.00%');
  } catch(_){}
}

/* ========= Metas / variações / indicadores ========= */

function _preencherMetaMensalAuto_({overwrite=false}={}) {
  // tenta Projeções centralizadas
  try { preverProximosMeses_(6, 'median'); } catch (_) {}

  try {
    metasUsarProjecaoSmart_({ overwrite });
    return;
  } catch (_) {}

  const prev = getSheetSmart_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']);
  const lanc = _getLanc_(), cfg = _getCfg_(); if (!prev || !lanc || !cfg) return;

  const lastPrev = prev.getLastRow(); if (lastPrev < 2) return;
  const itens = prev.getRange(2,1,lastPrev-1,1).getValues().flat();

  const today=_today_(), tz=_tz_();
  const mesesAlvo=[1,2,3].map(k=>new Date(today.getFullYear(), today.getMonth()-k, 1));
  const key=(d)=>Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(),1), tz, 'MM/yyyy');

  // Normalização unificada de Detalhamentos oficiais (case/acento)
  const lastCfg = _cfgLastRow_(cfg);
  const detOfRaw = (lastCfg>2 ? cfg.getRange(3,5,lastCfg-2,1).getValues() : []).flat().filter(Boolean);
  const detsOfNorm = new Set(detOfRaw.map(x => _normLower_(_stripDiacritics_(String(x)))));
  const norm = s => _normLower_(_stripDiacritics_(String(s||'')));

  const somaPorItemMes=new Map();
  const lastLanc=lanc.getLastRow();
  if (lastLanc >= 4) {
    const rows = lanc.getRange(4, COL.DATA, lastLanc-3, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues();
    const compMode=getCompMode_(cfg);

    for (const r of rows) {
      const [dataRaw, sub, det, , forma, parcelasRaw, valorRaw, valorParcRaw] = r;

      const detStr = String(det||'').trim();
      const subStr = String(sub||'').trim();
      const item   = detsOfNorm.has(norm(detStr)) ? detStr : subStr;
      if (!item) continue;

      const dt=(dataRaw instanceof Date && !isNaN(dataRaw)) ? dataRaw : parseDateBR_(dataRaw); if (!dt) continue;

      if (compMode === 'CONSUMO') {
        const k=key(dt);
        if (mesesAlvo.some(m=>key(m)===k)) {
          const tot=Number(valorRaw)||0;
          if (tot>0) {
            const kk=item+'|'+k;
            somaPorItemMes.set(kk,(somaPorItemMes.get(kk)||0)+tot);
          }
        }
        continue;
      }

      // CAIXA
      const cartRow=_findCartaoRow_(forma);
      let n=parseInt(parcelasRaw,10); if(!isFinite(n)||n<=0) n=1;
      const total=Number(valorRaw)||0; const parts=_parcelasExatas_(total, n);

      if (!cartRow) {
        // Modelo B não-cartão: VALOR=0 e VALOR_PARCELADO>0 → soma n*parcela
        const k=key(dt);
        if (mesesAlvo.some(m=>key(m)===k)) {
          const parc = Number(valorParcRaw)||0;
          const addVal = (total===0 && parc!==0) ? (parc*n) : total;
          if (addVal>0) {
            const kk=item+'|'+k;
            somaPorItemMes.set(kk,(somaPorItemMes.get(kk)||0)+addVal);
          }
        }
      } else {
        const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]);
        if (venc!=null && ini!=null && fim!=null) {
          for (let p=0;p<n;p++) {
            const dParc=new Date(dt); dParc.setMonth(dParc.getMonth()+p);
            const {mes,ano}=calcularCicloCartao(dParc, ini, fim, venc);
            const ref=new Date(ano, mes-1, 1);
            const k=key(ref);
            if (mesesAlvo.some(m=>key(m)===k)) {
              const v = _parcelaOuFallback_(valorParcRaw, parts[p]);
              if (_eqNum_(v, 0)) continue;
              const kk=item+'|'+k;
              somaPorItemMes.set(kk,(somaPorItemMes.get(kk)||0)+v);
            }
          }
        }
      }
    }
  }

  // Mediana dos 3 meses-alvo por item
  const metas = itens.map(item=>{
    const vals = mesesAlvo.map(m=>somaPorItemMes.get(item+'|'+key(m))||0).filter(v=>v>0);
    if(!vals.length) return [0];
    const sorted=vals.slice().sort((a,b)=>a-b);
    const mid=Math.floor(sorted.length/2);
    const med = sorted.length%2 ? sorted[mid] : (sorted[mid-1]+sorted[mid])/2;
    return [_r2(med)];
  });

  prev.getRange(2,3,metas.length,1).setValues(metas);
  try { prev.getRange(2,3,metas.length,1).setNumberFormat('R$ #,##0.00'); } catch(_){}
}

function _preencherColunaF_(){
  const prev=getSheetSmart_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']);
  const lanc=_getLanc_(), cfg=_getCfg_(); if(!prev||!lanc||!cfg) return;

  const lastPrev=prev.getLastRow(); if(lastPrev<2) return;
  const itens=prev.getRange(2,1,lastPrev-1,1).getValues().flat();
  const metas=prev.getRange(2,3,lastPrev-1,1).getValues().flat().map(Number);

  const today=_today_(), tz=_tz_();
  const alvoKey=Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), 1), tz, 'MM/yyyy');

  // Normalização de detalhamentos oficiais
  const lastCfg=_cfgLastRow_(cfg);
  const detOfRaw=(lastCfg>2? cfg.getRange(3,5,lastCfg-2,1).getValues():[]).flat().filter(Boolean);
  const detsOficiaisNorm=new Set(detOfRaw.map(x=>_normLower_(_stripDiacritics_(String(x)))));
  const norm=s=>_normLower_(_stripDiacritics_(String(s||'')));

  const sumMes=new Map();
  const lastLanc=lanc.getLastRow();
  if(lastLanc>=4){
    const rows=lanc.getRange(4, COL.DATA, lastLanc-3, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues();
    const compMode=getCompMode_(cfg);

    for(const r of rows){
      const [dataRaw, sub, det, , forma, parcelasRaw, valorRaw, valorParcRaw]=r;

      const detStr=String(det||'').trim();
      const subStr=String(sub||'').trim();
      const item = detsOficiaisNorm.has(norm(detStr)) ? detStr : subStr;
      if(!item) continue;

      const dt=(dataRaw instanceof Date&&!isNaN(dataRaw))?dataRaw:parseDateBR_(dataRaw);
      if(!dt) continue;

      if(compMode==='CONSUMO'){
        const k=Utilities.formatDate(new Date(dt.getFullYear(),dt.getMonth(),1), tz, 'MM/yyyy');
        if(k===alvoKey){
          const tot=Number(valorRaw)||0;
          if(tot>0) sumMes.set(item,(sumMes.get(item)||0)+tot);
        }
      } else {
        const cartRow=_findCartaoRow_(forma);
        let n=parseInt(parcelasRaw,10); if(!isFinite(n)||n<=0) n=1;

        if(!cartRow){
          // Modelo B não-cartão: VALOR=0 e VALOR_PARCELADO>0 → soma n*parcela no mês-alvo
          const k=Utilities.formatDate(new Date(dt.getFullYear(),dt.getMonth(),1), tz, 'MM/yyyy');
          if(k===alvoKey){
            const total=Number(valorRaw)||0;
            const parc = Number(valorParcRaw)||0;
            if (total===0 && parc!==0){
              sumMes.set(item,(sumMes.get(item)||0) + (parc*n));
            } else if (total>0){
              sumMes.set(item,(sumMes.get(item)||0) + total);
            }
          }
        } else {
          const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]);
          if(venc!=null&&ini!=null&&fim!=null){
            const total=Number(valorRaw)||0;
            // ✅ usa split exato em centavos
            const parts = _parcelasExatas_(total, n);
            for(let p=0;p<n;p++){
              const dParc=new Date(dt); dParc.setMonth(dParc.getMonth()+p);
              const {mes,ano}=calcularCicloCartao(dParc, ini, fim, venc);
              const k=Utilities.formatDate(new Date(ano,mes-1,1), tz, 'MM/yyyy');
              if(k===alvoKey){
                const v=_parcelaOuFallback_(valorParcRaw, parts[p]);
                if(_eqNum_(v,0)) continue;
                sumMes.set(item,(sumMes.get(item)||0)+v);
              }
            }
          }
        }
      }
    }
  }

  const out=itens.map((it,i)=>{
    const meta=Number(metas[i]||0);
    const gasto=Number(sumMes.get(String(it||'').trim())||0);
    if(!(meta>0)) return [''];
    return [(gasto/meta)-1];
  });

  prev.getRange(2,6,out.length,1).setValues(out);
  try{ prev.getRange(2,6,out.length,1).setNumberFormat('0.00%'); }catch(_){}
}


function _preencherIndicadorG_(){
  const sh=getSheetSmart_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']); if(!sh) return;
  const last=sh.getLastRow(); if(last<2) return;
  const varYTD = sh.getRange(2,5,last-1,1).getValues().flat().map(Number); // E
  const varMes = sh.getRange(2,6,last-1,1).getValues().flat().map(Number); // F

  const rot = (ytd, mes)=>{
    if(Number(mes)>0.15 || Number(ytd)>0.10) return 'Alta';
    if(Number(mes)>0.00 || Number(ytd)>0.00) return 'Média';
    return 'Baixa';
  };
  const out = varYTD.map((y,i)=>[ rot(y, varMes[i]) ]);
  sh.getRange(2,7,out.length,1).setValues(out);
}

function atualizarPrevisaoCompleta_({overwriteMetas=false}={}){
  ensureOrcamentoEstrutura_();   // A..G cabeçalho/formatos
  preencherPrevisaoDeGastos();   // A
  preencherColunaB();            // B (YTD por Item)
  _preencherMetaMensalAuto_({overwrite: overwriteMetas}); // C (de Projeções)
  atualizarOrcamentoVariacao_(); // D e E (usa B e C)
  _preencherColunaF_();          // F (% mês vs meta)
  _preencherIndicadorG_();       // G (Baixa/Média/Alta)
  _maybeToast_('✅ Previsão (A:G) atualizada.');
}

function atualizarOrcamentoVariacao_(){
  const sh=getSheetSmart_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']); if(!sh) return;
  const last=sh.getLastRow(); if(last<2) return;
  const realizado=sh.getRange(2,2,last-1,1).getValues().flat().map(Number);
  const metaMensal=sh.getRange(2,3,last-1,1).getValues().flat().map(Number);
  const varR=[], varP=[];
  for(let i=0;i<realizado.length;i++){
    const real=Number(realizado[i])||0, meta=Number(metaMensal[i])||0, metaAno=meta*12, d=real-metaAno;
    varR.push([_r2(d)]); varP.push([metaAno>0?((real/metaAno)-1):""]);
  }
  const rgR=sh.getRange(2,4,varR.length,1), rgP=sh.getRange(2,5,varP.length,1);
  rgR.setValues(varR); rgP.setValues(varP);
  try{ rgR.setNumberFormat('R$ #,##0.00'); rgP.setNumberFormat('0.00%'); }catch(_){}
}

/* ===================== PROJEÇÕES (median/winsor) ===================== */

function _median_(arr){
  const v=arr.slice().sort((a,b)=>a-b); const n=v.length;
  if(!n) return 0;
  const mid=Math.floor(n/2);
  return n%2? v[mid] : (v[mid-1]+v[mid])/2;
}

function preverProximosMeses_(nMeses, metodo){
  nMeses = Math.max(1, parseInt(nMeses,10)||3);
  metodo = String(metodo||'median').toLowerCase(); // 'mean' | 'median'

  const lanc = _getLanc_(), cfg = _getCfg_();
  if (!lanc || !cfg) return;

  const compMode = getCompMode_(cfg);
  const hoje = _today_(), tz = _tz_(), last = lanc.getLastRow();
  if (last < 4) return;

  const rows = lanc.getRange(4, COL.DATA, last-3, (COL.COMPETENCIA - COL.DATA + 1)).getValues();

  const soma = new Map();
  const mesKey = (d)=>Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), 1), tz, 'MM/yyyy');
  function add(k,v){ if(k && v>0) soma.set(k,(soma.get(k)||0)+Number(v)); }

  const mesCorrente = new Date(hoje.getFullYear(), hoje.getMonth(), 1).getTime();

  for (let i=0; i<rows.length; i++){
    const [dataRaw, subRaw, , , forma, parcelasRaw, valorRaw, valorParcRaw, /*status*/, compStr] = rows[i];
    const sub = String(subRaw||'').trim(); if(!sub) continue;
    const isCartao = ehCartao_(forma);

    if (compMode === 'CAIXA'){
      if (isCartao){
        const d=(dataRaw instanceof Date && !isNaN(dataRaw)) ? dataRaw : parseDateBR_(dataRaw); if(!d) continue;
        let n=parseInt(parcelasRaw,10); if(!isFinite(n)||n<=0) n=1;
        const total=Number(valorRaw)||0; const parts=_parcelasExatas_(total, n);
        const cartRow=_findCartaoRow_(forma); if(!cartRow) continue;
        const venc=_sanitizaDia_(cartRow[1]), ini=_sanitizaDia_(cartRow[2]), fim=_sanitizaDia_(cartRow[3]); if(venc==null||ini==null||fim==null) continue;

        for(let p=0;p<n;p++){
          const dataParc=new Date(d); dataParc.setMonth(dataParc.getMonth()+p);
          const { mes:mFat, ano:aFat }=calcularCicloCartao(dataParc, ini, fim, venc);
          const base=new Date(aFat, mFat-1, 1); if(base.getTime()>=mesCorrente) continue;
          const k=sub+'|'+mesKey(base);
          const v=_parcelaOuFallback_(valorParcRaw, parts[p]);
          if(_eqNum_(v,0)) continue;
          add(k, v);
        }
      } else {
        const d=(dataRaw instanceof Date && !isNaN(dataRaw)) ? dataRaw : parseDateBR_(dataRaw); if(!d) continue;
        const base=new Date(d.getFullYear(), d.getMonth(), 1); if(base.getTime()>=mesCorrente) continue;
        const v=Number(valorRaw)||0; if(v>0) add(sub+'|'+mesKey(base), v);
      }
    } else {
      // CONSUMO
      const d=(dataRaw instanceof Date && !isNaN(dataRaw)) ? dataRaw : parseDateBR_(dataRaw); if(!d) continue;
      const baseDate=new Date(d.getFullYear(), d.getMonth(), 1); if(baseDate.getTime()>=mesCorrente) continue;
      const v=Number(valorRaw)||0; if(v>0) add(sub+'|'+mesKey(baseDate), v);
    }
  }

  // agrega por subcategoria e projeta (median/mean com winsor)
  const agoraRef=new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  const subcats=new Set(Array.from(soma.keys()).map(k=>k.split('|')[0]));
  const proj=[];
  subcats.forEach(sub=>{
    const valores=[];
    for(let back=1; back<=24 && valores.length<12; back++){
      const d=new Date(agoraRef.getFullYear(), agoraRef.getMonth()-back, 1);
      const k=sub+'|'+mesKey(d);
      if(soma.has(k)) valores.push(soma.get(k));
    }
    if(valores.length){
      let estimativa;
      if(metodo==='mean'){
        estimativa = _r2(valores.reduce((a,b)=>a+b,0)/valores.length);
      } else {
        const v=valores.slice().sort((a,b)=>a-b), n=v.length;
        const lo=v[Math.floor(0.025*(n-1))], hi=v[Math.ceil(0.975*(n-1))];
        const clamp=v.map(x=>Math.min(Math.max(x,lo),hi));
        const mid=Math.floor(clamp.length/2);
        estimativa = clamp.length%2 ? clamp[mid] : _r2((clamp[mid-1]+clamp[mid])/2);
        estimativa = _r2(estimativa);
      }
      proj.push([sub, estimativa]);
    }
  });

  // ---------- Escrita idempotente (FIX dos “.” perdidos) ----------
  const mesesHeaders = Array.from({length:nMeses}, (_,i)=>{
    const d=new Date(hoje.getFullYear(), hoje.getMonth()+i+1, 1);
    return Utilities.formatDate(d, _tz_(), "MM/yyyy");
  });
  const sh = _safeEnsureSheet_('Projeções', ['Projecoes']);

  // FIX: usar spreads corretos
  const header = ['Subcategoria', `Previsão Mensal (${metodo})`, ...mesesHeaders];
  const nCols  = header.length;
  const linhas = proj.map(([sub, est]) => [sub, est, ...Array(nMeses).fill(est)]);
  const nRows  = Math.max(0, linhas.length);

  sh.getRange(1,1,1,nCols).setValues([header]);

  const prevLastRow = sh.getLastRow(), prevLastCol = sh.getLastColumn();
  if (nRows) sh.getRange(2,1,nRows,nCols).setValues(linhas);
  if (prevLastRow > nRows + 1) sh.getRange(nRows+2, 1, prevLastRow - (nRows+1), Math.max(nCols,1)).clearContent();
  if (prevLastCol > nCols)     sh.getRange(1, nCols+1, Math.max(1, prevLastRow), prevLastCol - nCols).clearContent();

  try{
    sh.getRange(1,1,1,nCols).setFontWeight('bold').setHorizontalAlignment('left');
    sh.setFrozenRows(1);
    for (let c=2; c<=nCols; c++){
      sh.getRange(2,c,Math.max(1, Math.max(sh.getLastRow()-1,1)),1).setNumberFormat('R$ #,##0.00');
    }
    try{ sh.setColumnWidth(1,260); for(let c=2;c<=nCols;c++) sh.setColumnWidth(c,140); }catch(_){}
  }catch(_){}

  _maybeToast_('📈 Projeções atualizadas ('+proj.length+' itens).');
}


function preverProximos3_(){ try{ preverProximosMeses_(3,'median'); }catch(e){ _maybeToast_('⚠️ Erro: '+(e&&e.message?e.message:e)); } }
function preverProximos6_(){ try{ preverProximosMeses_(6,'median'); }catch(e){ _maybeToast_('⚠️ Erro: '+(e&&e.message?e.message:e)); } }
function preverProximos12_(){ try{ preverProximosMeses_(12,'median'); }catch(e){ _maybeToast_('⚠️ Erro: '+(e&&e.message?e.message:e)); } }

/* ===================== AUDITORIAS ===================== */

function auditarDuplicados_(){ try{
  const sh=_getLanc_(); if(!sh) return;
  const last=sh.getLastRow(); if(last<4) return;
  const ids=sh.getRange(4, COL.ID_EXTRATO, last-3,1).getValues().flat(); const mapa=new Map();
  ids.forEach((id,i)=>{ if(!id) return; const row=i+4; const arr=mapa.get(id)||[]; arr.push(row); mapa.set(id,arr); });
  const rngAll=sh.getRange(4, COL.ID_EXTRATO, last-3,1); rngAll.setBackground(null).clearNote();
  mapa.forEach((linhas)=>{ if(linhas.length<=1) return; linhas.forEach(r=>{ const cell=sh.getRange(r, COL.ID_EXTRATO); cell.setBackground('#FFF3CD').setNote('Possível duplicado. Também em: '+linhas.join(', ')); }); });
}catch(e){ Logger.log('auditarDuplicados_ fast error: '+e); }}

function validarCartoes_(){
  const cfg=_getCfg_(); if(!cfg) return;
  const rows=_rangeCartoes_(cfg);
  const problemasCiclo=[], semLimite=[];
  rows.forEach(r=>{
    const nome=String(r[0]||"").trim();
    const venc=_sanitizaDia_(r[1]), ini=_sanitizaDia_(r[2]), fim=_sanitizaDia_(r[3]), lim=Number(r[4]||0);
    if(!nome) return;
    if(venc==null||ini==null||fim==null) problemasCiclo.push(`${nome}: ciclo inválido (início/fim/venc)`);
    if(!(lim>0)) semLimite.push(nome);
  });

  if(problemasCiclo.length){
    const max=12;
    const body=problemasCiclo.slice(0,max).join('\n')+(problemasCiclo.length>max?`\n+${problemasCiclo.length-max}…`:``);
    SpreadsheetApp.getUi().alert('Cartões com configuração inconsistente:\n'+body);
  } else {
    _maybeToast_('✔ Ciclos válidos.');
  }

  if(semLimite.length){
    const max=8;
    const msg = semLimite.length<=max ? semLimite.join(', ') : (semLimite.slice(0,max).join(', ') + ` +${semLimite.length-max}`);
    _maybeToast_('ℹ️ Cartões sem limite definido: ' + msg);
  }
}

function _alertarErrosResumo_(){
  const shFat=_ensureFaturasHeaders_(), shRes=_ensureResumoHeaders_(), ui=SpreadsheetApp.getUi(), tz=_tz_();
  if(!shFat||!shRes){ ui.alert('Auditoria', '❌ Abas "Faturas de Cartão" e/ou "Resumo de Faturas" não encontradas.', ui.ButtonSet.OK); return; }

  // Faturas -> esperado (líquido/pagos) por CHAVE
  const lastF=shFat.getLastRow(), fatRows=(lastF>=2)? shFat.getRange(2,1,lastF-1,11).getValues() : [], esperado=new Map();
  function addExp(key, liquido, pagos){ if(!esperado.has(key)) esperado.set(key,{liquido:0,pagos:0}); const o=esperado.get(key); o.liquido+=liquido; o.pagos+=pagos; }
  for(const r of fatRows){
    const cartao=String(r[0]||"").trim();
    const mesVal=r[1];
    const mesDate=(mesVal instanceof Date&&!isNaN(mesVal))? new Date(mesVal.getFullYear(), mesVal.getMonth(), 1):null;
    const mesStr = mesDate
      ? Utilities.formatDate(mesDate, tz, "MM/yyyy")
      : String(mesVal||"").trim();
    const total=Number(r[3]||0), pagos=Number(r[7]||0), encargos=Number(r[8]||0), creditos=Number(r[9]||0);
    if(!(cartao&&mesStr)) continue;
    const key=`${_cardKeyFromName_(cartao)}||${mesStr}`;
    const liquido=_r2(total+encargos-creditos);
    addExp(key, liquido, Math.max(0,pagos));
  }

  // Resumo -> atual por CHAVE
  const lastR=shRes.getLastRow(), resRows=(lastR>=2)? shRes.getRange(2,1,lastR-1,9).getValues() : [], atual=new Map();
  for(const r of resRows){
    const cartao=String(r[0]||"").trim();
       const mesVal=r[1];
    const mesDate=(mesVal instanceof Date&&!isNaN(mesVal))? new Date(mesVal.getFullYear(), mesVal.getMonth(), 1):null;
    const mesStr = mesDate
      ? Utilities.formatDate(mesDate, tz, "MM/yyyy")
      : "";
    const liquido=Number(r[2]||0), pend=Number(r[3]||0), conc=Number(r[4]||0), keyCol=String(r[6]||"").trim();
    let key=keyCol;
    if(!key){
      if(!cartao||!mesStr) continue;
      key=`${_cardKeyFromName_(cartao)}||${mesStr}`;
    }
    atual.set(key, { liquido:_r2(liquido), pend:_r2(pend), conc:_r2(conc) });
  }

  // Diferenças
  const diffs=[], faltantes=[], sobrando=[];
  function fmt(key){ const [cardHash, mes]=key.split('||'); return `${cardHash} — ${mes}`; }
  function neq(a,b){ return Math.abs((Number(a)||0)-(Number(b)||0))>=EPS_CENT; }

  for(const [key, exp] of esperado.entries()){
    if(!atual.has(key)){ faltantes.push(`+ Faltando no Resumo: ${fmt(key)} (líquido R$ ${exp.liquido.toFixed(2)})`); continue; }
    const cur=atual.get(key), campos=[];
    if(neq(exp.liquido,cur.liquido)) campos.push(`Líquido exp=${exp.liquido.toFixed(2)} ≠ cur=${cur.liquido.toFixed(2)}`);
    const pendExp=_r2(Math.max(exp.liquido-exp.pagos,0)), concExp=_r2(Math.min(exp.pagos,exp.liquido));
    if(neq(pendExp,cur.pend)) campos.push(`Pend exp=${pendExp.toFixed(2)} ≠ cur=${cur.pend.toFixed(2)}`);
    if(neq(concExp,cur.conc)) campos.push(`Conc exp=${concExp.toFixed(2)} ≠ cur=${cur.conc.toFixed(2)}`);
    if(campos.length) diffs.push(`• ${fmt(key)} → ${campos.join(' | ')}`);
  }
  for(const key of atual.keys()){ if(!esperado.has(key)) sobrando.push(`- Sobrando no Resumo: ${fmt(key)}`); }

  if(!diffs.length && !faltantes.length && !sobrando.length){
    SpreadsheetApp.getUi().alert('Auditoria', '✅ Resumo compatível com Faturas (pagto parcial).', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const cap=35, seg=(arr,t)=>arr.length?`\n${t}\n`+arr.slice(0,cap).join('\n')+(arr.length>cap?`\n... (+${arr.length-cap})`:``):'';
  const msg= seg(diffs,'Diferenças') + seg(faltantes,'Faltando no Resumo') + seg(sobrando,'Sobrando no Resumo');
  SpreadsheetApp.getUi().alert('Auditoria — Faturas × Resumo', msg.trim()||'Inconsistências detectadas.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function auditarGeralParte2_(){
  const ui = SpreadsheetApp.getUi();
  const tz = _tz_();
  const cfg = _getCfg_();
  const lanc = _getLanc_();
  const shFat = _ensureFaturasHeaders_();
  const shRes = _ensureResumoHeaders_();

  const linhas = [];

  // 1) Cartões
  try{
    const rows = _rangeCartoes_(cfg);
    const problemas=[], semLimite=[], duplicatas=[];
    const seen=new Set();
    rows.forEach(r=>{
      const nome=String(r[0]||"").trim();
      if(!nome) return;
      const k=_normLower_(_stripDiacritics_(nome));
      if(seen.has(k)) duplicatas.push(nome); else seen.add(k);

      const venc=_sanitizaDia_(r[1]), ini=_sanitizaDia_(r[2]), fim=_sanitizaDia_(r[3]), lim=Number(r[4]||0);
      if(venc==null||ini==null||fim==null) problemas.push(`${nome}: ciclo inválido (início/fim/venc)`);
      if(!(lim>0)) semLimite.push(nome);
    });
    if(problemas.length) linhas.push('• Cartões com ciclo inválido: '+(problemas.length<=6?problemas.join('; '): (problemas.slice(0,6).join('; ')+' …')));
    if(semLimite.length) linhas.push('• Cartões sem limite: '+(semLimite.length<=8?semLimite.join('; '):(semLimite.slice(0,8).join('; ')+' …')));
    if(duplicatas.length) linhas.push('• Cartões duplicados (nome): '+(duplicatas.length<=8?duplicatas.join('; '):(duplicatas.slice(0,8).join('; ')+' …')));
  }catch(_){}

  // 2) Faturas
  try{
    const last=shFat.getLastRow();
    const rows=(last>=2)? shFat.getRange(2,1,last-1,11).getValues():[];
    const keys=[], negCols=[], overPays=[];
    const map=new Map();
    rows.forEach(r=>{
      const cartao=String(r[0]||"").trim();
      const mesVal=r[1];
      const mesDate=(mesVal instanceof Date&&!isNaN(mesVal))? new Date(mesVal.getFullYear(), mesVal.getMonth(), 1):null;
      const mesStr = mesDate
        ? Utilities.formatDate(mesDate, tz, "MM/yyyy")
        : String(mesVal||"").trim();
      const val = Number(r[3]||0);
      const pago= Math.max(0, Number(r[7]||0));
      const enc = Math.max(0, Number(r[8]||0));
      const cred= Math.max(0, Number(r[9]||0));
      const chave=String(r[6]||"").trim() || (cartao&&mesStr ? `${_cardKeyFromName_(cartao)}||${mesStr}` : '');

      if(!chave) keys.push('(linha sem cartão/mês)');
      else map.set(chave, (map.get(chave)||0)+1);

      if(Number(r[7]||0) < 0 || Number(r[8]||0) < 0 || Number(r[9]||0) < 0){
        negCols.push(chave||'(sem chave)');
      }
      const liquido = _r2(val + enc - cred);
      if(pago > liquido + EPS_CENT) overPays.push(`${chave||'(sem chave)'} pago>${liquido.toFixed(2)}`);
    });
    const dups = Array.from(map.entries()).filter(([,n])=>n>1).map(([k])=>k);
    if(keys.length) linhas.push(`• Faturas sem CHAVE: ${keys.length}`);
    if(dups.length) linhas.push(`• Faturas com CHAVE duplicada: ${dups.length}`);
    if(negCols.length) linhas.push(`• Faturas com valores negativos em Pago/Encargos/Créditos: ${negCols.length}`);
    if(overPays.length) linhas.push(`• Faturas com pagamento acima do líquido: ${overPays.length}`);
  }catch(_){}

  // 3) Resumo
  try{
    const lastR=shRes.getLastRow();
    const rowsR=(lastR>=2)? shRes.getRange(2,1,lastR-1,9).getValues():[];
    const keysR=[], dupsR=[];
    const m=new Map();
    rowsR.forEach(r=>{
      const chave=String(r[6]||"").trim();
      if(!chave) keysR.push('(linha sem chave)');
      else m.set(chave,(m.get(chave)||0)+1);
    });
    dupsR.push(...Array.from(m.entries()).filter(([,n])=>n>1).map(([k])=>k));
    if(keysR.length) linhas.push(`• Resumo sem CHAVE: ${keysR.length}`);
    if(dupsR.length) linhas.push(`• Resumo com CHAVE duplicada: ${dupsR.length}`);

    // Divergências Resumo × Faturas
    const tz2=_tz_();
    const lastF=shFat.getLastRow();
    const fatRows=(lastF>=2)? shFat.getRange(2,1,lastF-1,11).getValues():[];
    const esperado=new Map();
    function addExp(key, liquido, pagos){ if(!esperado.has(key)) esperado.set(key,{liquido:0,pagos:0}); const o=esperado.get(key); o.liquido+=liquido; o.pagos+=pagos; }
    for(const r of fatRows){
      const cartao=String(r[0]||"").trim();
      const mesVal=r[1];
      const mesDate=(mesVal instanceof Date&&!isNaN(mesVal))? new Date(mesVal.getFullYear(), mesVal.getMonth(), 1):null;
      const mesStr = mesDate
        ? Utilities.formatDate(mesDate, tz2, "MM/yyyy")
        : String(mesVal||"").trim();
      const total=Number(r[3]||0), pagos=Number(r[7]||0), encargos=Number(r[8]||0), creditos=Number(r[9]||0);
      if(!(cartao&&mesStr)) continue;
      const key=`${_cardKeyFromName_(cartao)}||${mesStr}`;
      const liquido=_r2(total+encargos-creditos);
      addExp(key, liquido, Math.max(0,pagos));
    }
    const atual=new Map();
    rowsR.forEach(r=>{
      const cartao=String(r[0]||"").trim();
      const mesVal=r[1];
      const mesDate=(mesVal instanceof Date&&!isNaN(mesVal))? new Date(mesVal.getFullYear(), mesVal.getMonth(), 1):null;
      const mesStr = mesDate
        ? Utilities.formatDate(mesDate, tz2, "MM/yyyy")
        : "";
      const liquido=Number(r[2]||0), pend=Number(r[3]||0), conc=Number(r[4]||0), keyCol=String(r[6]||"").trim();
      let key=keyCol; if(!key){ if(!cartao||!mesStr) return; key=`${_cardKeyFromName_(cartao)}||${mesStr}`; }
      atual.set(key, { liquido:_r2(liquido), pend:_r2(pend), conc:_r2(conc), cartao });
    });
    let diffs=0, falt=0, sob=0;
    function neq(a,b){ return Math.abs((Number(a)||0)-(Number(b)||0))>=EPS_CENT; }
    for(const [key, exp] of esperado.entries()){
      if(!atual.has(key)){ falt++; continue; }
      const cur=atual.get(key);
      const pendExp=_r2(Math.max(exp.liquido-exp.pagos,0)), concExp=_r2(Math.min(exp.pagos,exp.liquido));
      if(neq(exp.liquido,cur.liquido) || neq(pendExp,cur.pend) || neq(concExp,cur.conc)) diffs++;
    }
    for(const key of atual.keys()){ if(!esperado.has(key)) sob++; }
    if(diffs||falt||sob) linhas.push(`• Resumo × Faturas: ${diffs} diferenças, ${falt} faltando, ${sob} sobrando`);

    // Utilização e Exposição
    const limites=_limitePorCartao_(cfg);
    const expoFutura=_exposicaoFuturaPorCartao_(lanc, cfg);
    let utilDiff=0, expoDiff=0;
    rowsR.forEach(r=>{
      const cartao=String(r[0]||"").trim(); if(!cartao) return;
      const liq=Number(r[2]||0);
      const utilCell=r[7]; // %
      const lim = Number(limites.get(cartao)||0);
      const utilEsperada = lim>0 ? liq/lim : '';
      const utilOk = (utilEsperada==='' && (utilCell===''||utilCell==null)) || (typeof utilEsperada==='number' && Math.abs(utilEsperada - Number(utilCell||0))<=0.005);
      if(!utilOk) utilDiff++;

      const expoCell=Number(r[8]||0);
      const expoEsperada=Number(expoFutura.get(cartao)||0);
      if(Math.abs(expoEsperada - expoCell) > 0.01) expoDiff++;
    });
    if(utilDiff) linhas.push(`• Resumo: Utilização (%) divergente em ${utilDiff} linha(s)`);
    if(expoDiff) linhas.push(`• Resumo: Exposição futura divergente em ${expoDiff} linha(s)`);
  }catch(_){}

  if(!linhas.length){
    ui.alert('Auditoria Geral (Parte 2)', '✅ Nenhuma inconsistência relevante encontrada.', ui.ButtonSet.OK);
  } else {
    const msg = 'Foram encontradas as seguintes pendências:\n\n- ' + linhas.join('\n- ');
    ui.alert('Auditoria Geral (Parte 2)', msg, ui.ButtonSet.OK);
  }
}

/** Gera lançamentos de demonstração (até 300 linhas) na aba Lançamento de Despesa. */
function gerar300LancamentosDemo_(qtd){
  qtd = Math.max(1, Math.min(Number(qtd)||300, 300));
  const sh = _getLanc_(), cfg = _getCfg_();
  if(!sh){ _maybeToast_('❌ Aba "Lançamento de Despesa" não encontrada.'); return; }

  // Garante que a DV da Forma está alinhada com Config (H + cartões)
  try{ corrigirDVFormaAgora_(); }catch(_){}

  const tz   = _tz_();
  const hoje = _today_();

  // Listas (com fallbacks)
  const cfgLists  = getCfgCached_(cfg) || {};
  const subs      = (cfgLists.subs && cfgLists.subs.length ? cfgLists.subs
                    : ['Mercado','Transporte','Saúde','Lazer','Restaurante','Serviços','Outros']);
  const subToCat  = cfgLists.subToCat || {};

  // Detalhamentos por Subcategoria
  let detPorSub = cfgLists.detPorSub;
  if (!detPorSub){
    detPorSub = {};
    try{
      const lastCfg = _cfgLastRow_(cfg);
      if (lastCfg && lastCfg >= 3){
        const detVals = cfg.getRange(3,5, lastCfg-2, 1).getValues().flat(); // E
        const subVals = cfg.getRange(3,6, lastCfg-2, 1).getValues().flat(); // F
        for (let i=0;i<detVals.length;i++){
          const det = String(detVals[i]||'').trim();
          const sub = String(subVals[i]||'').trim();
          if (det && sub){
            if (!detPorSub[sub]) detPorSub[sub] = [];
            detPorSub[sub].push(det);
          }
        }
      }
    }catch(_){}
  }

  // Formas de pagamento válidas
  let formasOK  = listaFormasComCartoes_(cfg) || [];
  if (!formasOK.length) formasOK = ['Pix','Débito','Dinheiro','Boleto'];

  // Helpers
  function rand(a,b){ return a + Math.random()*(b-a); }
  function pick(arr){ return arr[Math.floor(Math.random()*arr.length)]; }
  function randDateUltimosDias(maxDias){
    const d = new Date(hoje);
    d.setDate(d.getDate() - Math.floor(rand(0, maxDias)));
    return d;
  }
  function fmtComp(d){ return Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), 1), tz, 'MM/yyyy'); }

  const startRow      = Math.max(4, sh.getLastRow()+1);
  const rows          = [];
  const mesesAfetados = new Set();

  for(let i=0;i<qtd;i++){
    const d   = randDateUltimosDias(180);
    const sub = pick(subs);
    const cat = subToCat[sub] || 'Outros';

    let detalhamento = '';
    const detLista = detPorSub && detPorSub[sub];
    if (detLista && detLista.length) detalhamento = pick(detLista);

    const forma      = pick(formasOK);
    const usarCartao = ehCartao_(forma);

    let parcelas = 1;
    if (usarCartao && Math.random() < 0.35){
      parcelas = pick([2,3,4,6,8,10,12]);
    }

    let valor;
    if (/mercad|restaur|lazer|saud|transp|serv/i.test(sub)) {
      valor = _r2(rand(25, 600));
    } else {
      valor = _r2(rand(15, 1200));
    }

    const diasAteHoje = Math.floor((hoje - _dateOnly_(d))/(24*3600*1000));
    let status = '';
    let liquidacao = '';
    if (diasAteHoje > 45){
      status = 'Conciliado';
      liquidacao = d;
    } else if (diasAteHoje > 10){
      status = pick(['Pendente','Conciliado']);
      liquidacao = (status.match(/conciliado/i) ? d : '');
    } else {
      status = 'Pendente';
      liquidacao = '';
    }

    const comp = fmtComp(d);
    mesesAfetados.add(d.getMonth()+1);

    const idExtr = `DEMO-${Utilities.getUuid().slice(0,8)}-${(i+1)}`;

    const row = [];
    row[COL.DATA - COL.DATA]            = d;
    row[COL.SUBCATEGORIA - COL.DATA]    = sub;
    row[COL.DETALHAMENTO - COL.DATA]    = detalhamento;
    row[COL.CATEGORIA - COL.DATA]       = cat;
    row[COL.FORMA - COL.DATA]           = forma;
    row[COL.PARCELAS - COL.DATA]        = parcelas;
    row[COL.VALOR - COL.DATA]           = valor;
    row[COL.VALOR_PARCELADO - COL.DATA] = '';
    row[COL.STATUS - COL.DATA]          = status;
    row[COL.COMPETENCIA - COL.DATA]     = comp;
    row[COL.CENTRO_CUSTO - COL.DATA]    = '';
    row[COL.TIPO - COL.DATA]            = '';
    row[COL.LIQUIDACAO - COL.DATA]      = liquidacao;
    row[COL.ID_EXTRATO - COL.DATA]      = idExtr;
    row[COL2.COMP_CONSUMO - COL.DATA]   = comp;
    row[COL_FP - COL.DATA]              = '';

    rows.push(row);
  }

  // Grava em bloco
  sh.getRange(startRow, COL.DATA, rows.length, (COL_FP - COL.DATA + 1)).setValues(rows);

  // Formatos
  try{
    sh.getRange(startRow, COL.DATA, rows.length, 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(startRow, COL.VALOR, rows.length, 1).setNumberFormat('R$ #,##0.00');
    sh.getRange(startRow, COL.VALOR_PARCELADO, rows.length, 1).setNumberFormat('R$ #,##0.00');
    sh.getRange(startRow, COL.COMPETENCIA, rows.length, 1).setNumberFormat('MM/yyyy');
    sh.getRange(startRow, COL2.COMP_CONSUMO, rows.length, 1).setNumberFormat('MM/yyyy');
    sh.getRange(startRow, COL.LIQUIDACAO, rows.length, 1).setNumberFormat('dd/MM/yyyy');
  }catch(_){}

  try{ if (typeof pendMarcarMeses_ === 'function') pendMarcarMeses_([...mesesAfetados]); }catch(_){}
  try{ if (typeof pendMarcarPrevisao_ === 'function') pendMarcarPrevisao_(); }catch(_){}
  try{ if (typeof pendMarcarFaturasResumo_ === 'function') pendMarcarFaturasResumo_(); }catch(_){}

  _maybeToast_(`✅ Inseridos ${rows.length} lançamentos demo; Detalhamento só quando existir para a Subcategoria. Recalcule para atualizar resumos/indicadores.`);
}

/***************************************************************
 * FINANCEIRO FAMILIAR — SCRIPT UNIFICADO (PARTE 2 — BLOCO 2C)
 * Indicadores (KPI), Reset backend, Processador de Pendências,
 * Colunas/Formatos, Menus/Triggers (onOpen).
 ***************************************************************/

/** === SHIMS usados no 2C (criados só se NÃO existirem) === */

// Toast seguro (fallback)
if (typeof this._maybeToast_ !== 'function') {
  this._maybeToast_ = function(msg){
    try{ SpreadsheetApp.getActive().toast(String(msg||'')); }
    catch(_){ Logger.log(String(msg||'')); }
  };
}

// Lock reentrante simples (mesma execução)
var __DOCLOCK_HELD__ = (typeof __DOCLOCK_HELD__ === 'boolean') ? __DOCLOCK_HELD__ : false;
if (typeof this.withDocLock_ !== 'function') {
  this.withDocLock_ = function(tag, fn){
    if (typeof fn !== 'function') return null;

    if (__DOCLOCK_HELD__) {               // já estamos sob lock? executa direto
      try { return fn(); }
      catch(e){ _maybeToast_('⚠️ '+(tag||'processo')+': '+(e && e.message || e)); throw e; }
    }

    const lock = LockService.getDocumentLock();
    let ok = lock.tryLock(5000);
    for (let i=0; !ok && i<2; i++){ Utilities.sleep(400); ok = lock.tryLock(5000); }
    if (!ok){ _maybeToast_('⏳ '+(tag||'processo')+' em andamento…'); return; }

    try {
      __DOCLOCK_HELD__ = true;
      return fn();
    } finally {
      __DOCLOCK_HELD__ = false;
      try { lock.releaseLock(); } catch(_){}
    }
  };
}

// Logs básicos (tolerantes à ausência de LOG_)
if (typeof this._logInfo_ !== 'function') {
  this._logInfo_ = function(){ 
    try{ this.LOG_ && this.LOG_.info && this.LOG_.info.apply(null, arguments); }
    catch(_){ Logger.log('[INFO] '+[].slice.call(arguments).join(' ')); }
  };
}
if (typeof this._logWarn_ !== 'function') {
  this._logWarn_ = function(){ 
    try{ this.LOG_ && this.LOG_.warn && this.LOG_.warn.apply(null, arguments); }
    catch(_){ Logger.log('[WARN] '+[].slice.call(arguments).join(' ')); }
  };
}

// Planilha de Logs (fallback) — unificada para "LOGS"
if (typeof this._ensureLogsSheet_ !== 'function') {
  this._ensureLogsSheet_ = function(){
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('LOGS');
    if(!sh) sh = ss.insertSheet('LOGS');
    if (sh.getLastRow() === 0){
      sh.getRange(1,1,1,3).setValues([['Quando','Nível','Mensagem']]);
    }
    return sh;
  };
}

// Shims utilitários mínimos
if (typeof this._r2 !== 'function') {
  this._r2 = function(v){ return Math.round(Number(v||0)*100)/100; };
}
if (typeof this._dateOnly_ !== 'function') {
  this._dateOnly_ = function(d){ const x=new Date(d); x.setHours(0,0,0,0); return x; };
}
if (typeof this._isConciliadoStatus_ !== 'function') {
  this._isConciliadoStatus_ = function(s){ return /conciliad|quitad|liquidad|ok/i.test(String(s||'')); };
}
if (typeof this._parcelaOuFallback_ !== 'function') {
  this._parcelaOuFallback_ = function(valorParcRaw, fallbackCents){
    const n = Number(valorParcRaw);
    return isFinite(n) && n>0 ? n : fallbackCents;
  };
}

/* ==== Pendências com flags/meses persistidos (cache de script) ==== */
(function(global){
  const C = CacheService.getScriptCache();
  const K_FULL = 'PEND_FULL';
  const K_MESES = 'PEND_MESES';   // JSON: [1..12]
  const K_PREV  = 'PEND_PREV';    // '1' | ''
  const K_FAT   = 'PEND_FAT';     // '1' | ''
  const TTL = 1800; // 30 minutos (↑)

  if (typeof global.pendMarcarFull_ !== 'function')
    global.pendMarcarFull_ = function(){ try{ C.put(K_FULL,'1',TTL); }catch(_){ } };

  if (typeof global.pendMarcarMeses_ !== 'function')
    global.pendMarcarMeses_ = function(meses){
      try{
        const arr = (Array.isArray(meses)? meses: [])
          .map(n=>Math.max(1,Math.min(12,Number(n)||0))).filter(Boolean);
        if (arr.length) C.put(K_MESES, JSON.stringify(arr), TTL);
      }catch(_){}
    };

  if (typeof global.pendMarcarPrevisao_ !== 'function')
    global.pendMarcarPrevisao_ = function(){ try{ C.put(K_PREV,'1',TTL); }catch(_){ } };

  if (typeof global.pendMarcarFaturasResumo_ !== 'function')
    global.pendMarcarFaturasResumo_ = function(){ try{ C.put(K_FAT,'1',TTL); }catch(_){ } };

  if (typeof global.pendLerEApagar_ !== 'function')
    global.pendLerEApagar_ = function(){
      let full=false, meses=null, prev=false, fat=false;
      try{
        full = (C.get(K_FULL)==='1'); if(full) C.remove(K_FULL);
        prev = (C.get(K_PREV)==='1');       C.remove(K_PREV);
        fat  = (C.get(K_FAT)==='1');        C.remove(K_FAT);

        const raw = C.get(K_MESES);
        if (raw){ try{ meses = JSON.parse(raw)||null; }catch(_){ meses=null; } }
        C.remove(K_MESES);
      }catch(_){}

      if (full) meses = [1,2,3,4,5,6,7,8,9,10,11,12];
      if (!meses || !meses.length){
        const hoje = new Date(); meses = [ hoje.getMonth()+1 ];
      }
      return { full, meses, prev, fat };
    };
})(this);

/* ===================== INDICADORES (KPI) ===================== */
function ensureIndicadores_(){
  return withDocLock_('ensureIndicadores_', () => {
    const ss = SpreadsheetApp.getActive();
    const sh = getSheetSmart_(ABAS.INDICADORES) || ss.insertSheet(ABAS.INDICADORES);

    try { sh.clearContents(); } catch(_) { sh.clear(); }
    sh.getRange(1,1,1,2).setValues([['Indicador','Valor']]);

    const tz   = _tz_();
    const res  = _ensureResumoHeaders_();
    const lanc = _getLanc_();
    const cfg  = _getCfg_();

    const hoje        = _today_();
    const mesAtual    = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    const mesAtualStr = Utilities.formatDate(mesAtual, tz, "MM/yyyy");

    // Utilização média / máxima e Top-3 (texto)
    let utilMedia = '', utilMax = '', top3Str = '';
    try{
      if(res && cfg){
        const limites = _limitePorCartao_(cfg);
        const last    = res.getLastRow();
        const rows    = (last >= 2) ? res.getRange(2,1,last-1,9).getValues() : [];

        const liqPorCartao = new Map();
        rows.forEach(r=>{
          const cartao = String(r[0]||'').trim();
          const mesVal = r[1];
          const mesStr = (mesVal instanceof Date && !isNaN(mesVal))
            ? Utilities.formatDate(new Date(mesVal.getFullYear(), mesVal.getMonth(), 1), tz, "MM/yyyy")
            : String(mesVal||'').trim();
          if(!cartao || mesStr !== mesAtualStr) return;
          const liquido = Number(r[2]||0);
          if(liquido > 0) liqPorCartao.set(cartao, (liqPorCartao.get(cartao)||0) + liquido);
        });

        const cards = [];
        liqPorCartao.forEach((liq, cartao)=>{
          const lim = Number(limites.get(cartao)||0);
          if(lim > 0) cards.push({ cartao, util: liq/lim, liquido: liq, limite: lim });
        });

        if(cards.length){
          const utils = cards.map(c=>c.util);
          utilMedia = utils.reduce((a,b)=>a+b,0)/utils.length;
          utilMax   = Math.max.apply(null, utils);

          cards.sort((a,b)=> b.util - a.util);

          function fmtR(v){ return 'R$ ' + _r2(v).toFixed(2); }
          top3Str = cards.slice(0,3)
            .map(x => `${x.cartao} ${(x.util*100).toFixed(0)}% (${fmtR(x.liquido)} / ${fmtR(x.limite)})`)
            .join('; ');
        }
      }
    }catch(e){ _logWarn_('ensureIndicadores_', 'Falha ao calcular utilização', {err:e.message}); }

    // Exposição futura
    let expoTotal = 0, expoBuckets = { d30:0, d90:0, d180:0, dInf:0 };
    try{
      if(lanc && cfg){
        const perCard = _exposicaoFuturaPorCartao_(lanc, cfg);
        expoTotal = Array.from(perCard.values()).reduce((a,b)=>a+(Number(b)||0),0);
        expoBuckets = _exposicaoFuturaBuckets_(lanc, cfg);
      }
    }catch(e){ _logWarn_('ensureIndicadores_', 'Falha ao calcular exposição', {err:e.message}); }

    // Orçamento
    let desvTot = '', pctAcima = '';
    try{
      const prev = getSheetSmart_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']);
      if(prev){
        const last = prev.getLastRow();
        if(last >= 2){
          const varR  = prev.getRange(2,4,last-1,1).getValues().flat().map(Number);
          const varP  = prev.getRange(2,5,last-1,1).getValues().flat().map(Number);
          const metas = prev.getRange(2,3,last-1,1).getValues().flat().map(Number);
          if(varR.length){
            desvTot = varR.reduce((a,b)=>a+(Number(b)||0),0);
            const itensComMeta = metas.map(x=>x>0);
            const acima = varP.filter((v,i)=> itensComMeta[i] && Number(v)>0).length;
            const base  = itensComMeta.filter(Boolean).length;
            pctAcima = base>0 ? (acima/base) : '';
          }
        }
      }
    }catch(e){ _logWarn_('ensureIndicadores_', 'Falha orçamento', {err:e.message}); }

    // Taxa de conciliação
    let taxaConc = '';
    try{
      if(res){
        const lastR = res.getLastRow();
        const rowsR = (lastR>=2) ? res.getRange(2,1,lastR-1,9).getValues() : [];
        let liquidoTot=0, concTot=0;
        rowsR.forEach(r=>{
          const mesVal=r[1];
          const mesStr=(mesVal instanceof Date && !isNaN(mesVal))
            ? Utilities.formatDate(new Date(mesVal.getFullYear(), mesVal.getMonth(), 1), tz, "MM/yyyy")
            : String(mesVal||'').trim();
          if(mesStr!==mesAtualStr) return;
          liquidoTot += Number(r[2]||0);
          concTot    += Number(r[4]||0);
        });
        taxaConc = (liquidoTot>0) ? (concTot/liquidoTot) : '';
      }
    }catch(e){ _logWarn_('ensureIndicadores_', 'Falha conciliação', {err:e.message}); }

    const out = [
      ['Utilização média cartões (%)',                utilMedia || ''],
      ['Utilização máxima (pior cartão) (%)',         utilMax   || ''],
      ['Top-3 cartões por utilização',                top3Str   || ''],
      ['Exposição futura ≤30d (R$)',                  expoBuckets.d30 || ''],
      ['Exposição futura 31–90d (R$)',                expoBuckets.d90 || ''],
      ['Exposição futura 91–180d (R$)',               expoBuckets.d180 || ''],
      ['Exposição futura >180d (R$)',                 expoBuckets.dInf || ''],
      ['Exposição futura total (R$)',                 expoTotal || ''],
      ['Orçamento — Desvio total (R$)',               desvTot   || ''],
      ['Orçamento — % de itens acima da meta',        pctAcima  || ''],
      ['Taxa de conciliação (mês corrente)',          taxaConc  || ''],
      ['Atualizado em',                               new Date()]
    ];
    sh.getRange(2,1,out.length,2).setValues(out);

    try{
      // sh.getRange("B1").setNumberFormat('@'); // ← removido (ruído)
      sh.getRange("B2:B3").setNumberFormat('0.00%');
      sh.getRange("B4").setNumberFormat('@');
      sh.getRange("B5:B9").setNumberFormat('R$ #,##0.00');
      sh.getRange("B10").setNumberFormat('R$ #,##0.00');
      sh.getRange("B11").setNumberFormat('0.00%');
      sh.getRange("B12").setNumberFormat('0.00%');
      sh.getRange("B13").setNumberFormat('dd/MM/yyyy HH:mm');
    }catch(_){}

    _maybeToast_('✅ Indicadores (KPI) atualizados.');
    _logInfo_('ensureIndicadores_', 'concluído');
  });
}

/* Atualização manual via menu */
function atualizarIndicadores_(){
  try{
    ensureIndicadores_();
    _maybeToast_('✅ Indicadores (KPI) atualizados.');
  }catch(e){
    _maybeToast_('⚠️ Erro ao atualizar indicadores: ' + (e && e.message ? e.message : e));
  }
}

/* ===================== LOGS (trim) ===================== */
function _trimLogs_(){ 
  try { 
    const sh = _ensureLogsSheet_(); 
    const max = 2000; 
    const n = sh.getLastRow(); 
    if (n > max + 1) sh.deleteRows(2, n - max - 1); 
  } catch(_) {} 
}

/* ===================== PROCESSADOR ÚNICO (pendências) ===================== */
function processarPendencias_(){
  const lock=LockService.getDocumentLock();
  let ok=lock.tryLock(5000);
  for(let i=0; !ok && i<2; i++){ Utilities.sleep(400); ok=lock.tryLock(5000); }
  if(!ok){ _maybeToast_('⛳ Outra atualização em curso...'); return; }

  try{
    const shCfg=_getCfg_();
    const shLanc=_getLanc_();
    if(!shCfg||!shLanc) return;

    const pend=pendLerEApagar_(); // { full, meses, prev, fat }

    // 1) Mensal (por competência)
    if (pend.full || (pend.meses&&pend.meses.length)){
      const meses=pend.full ? [1,2,3,4,5,6,7,8,9,10,11,12] : pend.meses;
      meses.forEach(m=>{
        const shMes=getSheetSmart_(MESES[m-1]);
        if(shMes){ try{ atualizarResultadosMensaisComCartoes(shLanc, shMes, shCfg, m, 3, 14); }catch(e){ Logger.log(e); } }
      });
    }

    // 2) Previsão/Orçamento
    if (pend.prev){
      try{ preencherPrevisaoDeGastos(); }catch(_){}
      try{ preencherColunaB(); }catch(_){}
      try{ ensureOrcamentoEstrutura_(); }catch(_){}
      try{ atualizarOrcamentoVariacao_(); }catch(_){}
    }

    // 3) Faturas & Resumo — sem relock nas internas
    if (pend.fat){
      try{ gerarFaturasDeCartao_(false); }catch(_){}
      try{ atualizarResumoFaturas_(false); }catch(_){}
      try{ rebuildParcelasCartao_(false); }catch(_){}
      try{ sincronizarStatusLancamentosComResumo_({ downgrade:true }); }catch(_){}
    }

    try{ _trimLogs_(); }catch(_){}
    _maybeToast_('✅ Atualização concluída.');
  }catch(e){ _maybeToast_('⚠️ Erro ao processar: '+(e&&e.message?e.message:e)); }
  finally{ try{ lock.releaseLock(); }catch(_){ } }
}

/* ===================== COLUNAS TÉCNICAS (ocultar) & FORMATOS ===================== */
function _hideTechCols_(){
  const sh=_getLanc_(); if(!sh) return;
  try{ sh.hideColumns(COL2.COMP_CONSUMO); }catch(_){}
  try{ sh.hideColumns(COL_FP); }catch(_){}
}
function _ensureFormatsLanc_(){
  const sh=_getLanc_(); if(!sh) return;
  try{
    const n = Math.max(0, sh.getMaxRows()-3);
    sh.getRange(4, COL.LIQUIDACAO, n, 1).setNumberFormat("dd/MM/yyyy");
    sh.getRange(4, COL.COMPETENCIA, n, 1).setNumberFormat("MM/yyyy");
    sh.getRange(4, COL2.COMP_CONSUMO, n, 1).setNumberFormat("MM/yyyy");
  }catch(_){}
}

/***************************************************************
 * PARCELAS DO CARTÃO — nova aba
 * - Cria/estiliza a aba "Parcelas do Cartão" com padrão de cores
 * - Reconstrói a tabela “explodindo” os lançamentos parcelados
 * - Status da parcela = "Conciliado" quando a fatura (Resumo) está liquidada
 ***************************************************************/
function _ensureAbaParcelas_(){
  const ss = SpreadsheetApp.getActive();
  let sh = getSheetSmart_(ABAS.PARCELAS_CARTAO);
  if(!sh) sh = ss.insertSheet(ABAS.PARCELAS_CARTAO);

  // Paleta
  const HEADER_BG = '#9fc5e8';

  // 0) Garante que NÃO há merge na linha 1 (evita erro ao congelar)
  try { sh.getRange(1,1,1, sh.getMaxColumns()).breakApart(); } catch(_){}

  // 1) Título em F1 (sem merge) + wrap
  const titleCell = sh.getRange('F1');
  const curTitle  = String(titleCell.getValue() || '').trim();
  if (!curTitle) titleCell.setValue('Parcelas do Cartão');
  try {
    if (titleCell.setWrapStrategy) titleCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    else titleCell.setWrap(true);
  } catch(_) {}
  try { sh.setRowHeight(1, 28); } catch(_){}

  // 2) Cabeçalho na linha 2 (só cria/formata se não estiver pronto)
  const hdr = [
    'Cartão','Mês Fatura','Subcategoria','Detalhamento',
    'Data Compra','Parcela (n/N)','Valor Parcela (R$)',
    'Status Fatura','CHAVE Fatura','Categoria','Forma','ID/Tx'
  ];
  const rngHdr = sh.getRange(2,1,1,hdr.length);
  const curHdr = rngHdr.getValues()[0].map(v => String(v||'').trim());
  const headerIgual = curHdr.join('|') === hdr.join('|');
  if (!headerIgual) rngHdr.setValues([hdr]);

  // formatação do cabeçalho (sem brigar com usuário)
  try {
    const hasBg = rngHdr.getBackgrounds()[0].some(c => c.toLowerCase() !== '#ffffff');
    if (!hasBg) rngHdr.setFontFamily('Arial').setFontWeight('bold').setBackground(HEADER_BG);
  } catch(_) {}

  // 3) Congelar linhas
  try {
    if (sh.getFrozenRows() < 2) sh.setFrozenRows(2);
    if (sh.getFrozenColumns() > 0) sh.setFrozenColumns(0);
  } catch(_) {}

  // 4) Formatos de número
  try { sh.getRange('B:B').setNumberFormat('MM/yyyy'); } catch(_){}
  try { sh.getRange('E:E').setNumberFormat('dd/MM/yyyy'); } catch(_){}
  try { sh.getRange('G:G').setNumberFormat('R$ #,##0.00'); } catch(_){}

  // 4.1) Normaliza valores de B (1º dia do mês)
  try {
    const last = Math.max(3, sh.getLastRow());
    if (last >= 3) {
      const vals = sh.getRange(3,2,last-2,1).getValues();
      let needWrite = false;
      for (let i=0;i<vals.length;i++){
        const v = vals[i][0];
        if (v instanceof Date && !isNaN(v)) {
          const norm = new Date(v.getFullYear(), v.getMonth(), 1);
          if (+norm !== +v) { vals[i][0] = norm; needWrite = true; }
        }
      }
      if (needWrite){
        sh.getRange(3,2,vals.length,1).setValues(vals).setNumberFormat('MM/yyyy');
      }
    }
  } catch(_){}

  // 5) Larguras (heurística leve)
  try {
    const widths=[170,100,160,200,110,110,140,130,160,150,140,120];
    const curr = sh.getColumnWidth(1);
    if (curr < 160) { widths.forEach((w,i)=>{ try{ sh.setColumnWidth(i+1,w); }catch(_){ } }); }
  } catch(_) {}

  // 6) Proteções idempotentes
  try { _protectHeaderRow_(sh); } catch(_){}
  try{
    const DESC = 'Cabeçalho (linha 2) protegido';
    const prots = (sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)||[])
      .filter(p => (p.getDescription()||'') === DESC);
    const ok = prots.some(p => {
      try{
        const rg = p.getRange();
        return rg && rg.getRow()===2 && rg.getNumRows()===1 &&
               rg.getColumn()===1 && rg.getNumColumns()===sh.getMaxColumns();
      }catch(_){ return false; }
    });
    if (!ok){
      prots.forEach(p=>{ try{ p.remove(); }catch(_){ } });
      const pr = sh.getRange(2,1,1, Math.max(1, sh.getMaxColumns())).protect().setDescription(DESC);
      try{ pr.setWarningOnly(false); }catch(_){}
      try{ pr.removeEditors(pr.getEditors()); }catch(_){}
    }
  }catch(_){}

  return sh;
}

/** Reconstrói a aba com todas as parcelas (1 linha por parcela). */
function rebuildParcelasCartao_(){
  return withDocLock_('rebuildParcelasCartao_', () => {
    const shLanc = _getLanc_();
    const shCfg  = _getCfg_();
    const shOut  = _ensureAbaParcelas_();
    const shRes  = getSheetSmart_(ABAS.RESUMO_FATURAS);

    if (!shLanc || !shCfg || !shOut) {
      _maybeToast_('❌ Preciso das abas: Lançamento de Despesa, Configurações e Parcelas do Cartão.');
      _logWarn_('rebuildParcelasCartao_', 'Abas necessárias ausentes');
      return;
    }

    const tz = _tz_();

    // Cabeçalho mínimo na linha 2 (por garantia)
    const HEADERS = [
      'Cartão','Mês Fatura','Subcategoria','Detalhamento',
      'Data Compra','Parcela (n/N)','Valor Parcela (R$)',
      'Status Fatura','CHAVE Fatura','Categoria','Forma','ID/Tx'
    ];
    try{
      const lastC = Math.max(12, shOut.getLastColumn());
      const curHdr = shOut.getRange(2,1,1,lastC).getValues()[0];
      if (HEADERS.some((h,i)=> String(curHdr[i]||'').trim()!==h)){
        shOut.getRange(2,1,1,12).setValues([HEADERS]);
      }
    }catch(_){}

    // CHAVES de faturas liquidadas (Pendentes ≈ 0) no Resumo
    const paidKeys = new Set();
    try{
      if (shRes){
        const lastR = shRes.getLastRow();
        if (lastR >= 2){
          const rows = shRes.getRange(2,1,lastR-1,9).getValues();
          for (const r of rows){
            const cartao  = String(r[0]||'').trim();
            const mesVal  = r[1];
            const liquido = Number(r[2]||0);
            const pend    = Number(r[3]||0);
            if (!(cartao && mesVal)) continue;

            const mesStr = (mesVal instanceof Date && !isNaN(mesVal))
              ? Utilities.formatDate(new Date(mesVal.getFullYear(), mesVal.getMonth(), 1), tz, "MM/yyyy")
              : String(mesVal||'').trim();

            const key = (String(r[6]||'').trim()) || (_cardKeyFromName_(cartao)+'||'+mesStr);
            if (liquido > 0 && Math.abs(pend) <= EPS_CENT) paidKeys.add(key);
          }
        }
      }
    }catch(e){ _logWarn_('rebuildParcelasCartao_', 'Falha ao ler Resumo', {err: e.message}); }

    // Explode lançamentos de cartão em parcelas
    const out = [];
    const last = shLanc.getLastRow();
    if (last >= 4){
      const rows = shLanc.getRange(4, COL.DATA, last-3, (COL.VALOR_PARCELADO - COL.DATA + 1)).getValues();
      for (const r of rows){
        const [dataRaw, sub, det, cat, forma, parcelasRaw, valorRaw, valorParcRaw] = r;

        // Data da compra
        const dt = (dataRaw instanceof Date && !isNaN(dataRaw)) ? dataRaw : parseDateBR_(dataRaw);
        if (!dt) continue; // ignora datas inválidas

        // Apenas lançamentos com forma = cartão
        const cart = _findCartaoRow_(forma);
        if (!cart) continue;

        const cartaoNome = String(cart[0]||'').trim();
        const venc = _sanitizaDia_(cart[1]);
        const ini  = _sanitizaDia_(cart[2]);
        const fim  = _sanitizaDia_(cart[3]);
        if (venc==null || ini==null || fim==null) continue;

        // Parcelamento
        const total = Number(valorRaw)||0;
        let n = parseInt(parcelasRaw,10); if(!isFinite(n)||n<=0) n=1;
        const parts = _parcelasExatas_(total, n); // distribuição exata

        // Parcela a parcela, posicionada no mês/ano da FATURA
        for (let p=0; p<n; p++){
          const dParc = new Date(dt); dParc.setMonth(dParc.getMonth()+p);
          const { mes:mFat, ano:aFat } = calcularCicloCartao(dParc, ini, fim, venc);
          const mesRef = new Date(aFat, mFat-1, 1);
          const mesKey = Utilities.formatDate(mesRef, tz, 'MM/yyyy');

          // Valor da parcela: prioriza a coluna "Valor Parcela (I)" quando presente
          const vParc = _parcelaOuFallback_(valorParcRaw, parts[p]);
          if (_eqNum_(vParc, 0)) continue; // evita linhas "fantasma"

          const fKey  = _cardKeyFromName_(cartaoNome) + '||' + mesKey;
          const stFat = paidKeys.has(fKey) ? 'Conciliado' : 'Pendente';

          out.push([
            cartaoNome, mesRef, String(sub||''), String(det||''),
            _dateOnly_(dt), ((p+1)+'/'+n), _r2(vParc),
            stFat, fKey, String(cat||''), String(forma||''), ''
          ]);
        }
      }
    }

    // Limpa dados antigos e grava
    const DATA_START_ROW = 3;
    try{
      const lastOut = shOut.getLastRow();
      if (lastOut >= DATA_START_ROW) {
        const rowsToClear = lastOut - DATA_START_ROW + 1;
        shOut.getRange(DATA_START_ROW, 1, rowsToClear, 12).clearContent().clearNote();
      }
    }catch(_){}

    if (out.length){
      shOut.getRange(DATA_START_ROW, 1, out.length, 12).setValues(out);
    }

    _maybeToast_(`✅ Parcelas do Cartão atualizadas — ${out.length} linha(s).`);
    _logInfo_('rebuildParcelasCartao_', 'concluído', {linhas: out.length});
  });
}

/** Sobe "Conciliado" nos Lançamentos de CARTÃO quando TODAS as parcelas
 * caem em faturas já liquidadas no Resumo. Também preenche Data de Liquidação.
 * downgrade=false mantém Conciliado mesmo se a fatura voltar a Pendente.
 */
function sincronizarStatusLancamentosComResumo_({ downgrade = true } = {}) {
  return withDocLock_('sincronizarStatusLancamentosComResumo_', () => {
    const shLanc = _getLanc_();
    const shRes  = _ensureResumoHeaders_();
    const shCfg  = _getCfg_();
    if (!shLanc || !shRes || !shCfg) return;

    const tz = _tz_(), paidKeys = new Set();

    // CHAVES de faturas liquidadas no Resumo (Pendentes ≈ 0)
    const lastR = shRes.getLastRow();
    if (lastR >= 2) {
      const rows = shRes.getRange(2, 1, lastR - 1, 9).getValues();
      for (const r of rows) {
        const cartao  = String(r[0] || '').trim();
        const mesVal  = r[1];
        const liquido = Number(r[2] || 0);
        const pend    = Number(r[3] || 0);
        if (!(cartao && mesVal)) continue;

        const mesStr = (mesVal instanceof Date && !isNaN(mesVal))
          ? Utilities.formatDate(new Date(mesVal.getFullYear(), mesVal.getMonth(), 1), tz, "MM/yyyy")
          : String(mesVal || '').trim();

        const key = (String(r[6] || '').trim()) || (_cardKeyFromName_(cartao) + '||' + mesStr);
        if (liquido > 0 && Math.abs(pend) <= EPS_CENT) paidKeys.add(key);
      }
    }

    // Percorre Lançamentos (B..N)
    const first = 4, last = shLanc.getLastRow();
    if (last < first) return;

    const n = last - first + 1;
    const bloc = shLanc.getRange(first, COL.DATA, n, (COL.LIQUIDACAO - COL.DATA + 1)).getValues();
    const curStatusCol = shLanc.getRange(first, COL.STATUS, n, 1).getValues();
    const curLiqCol    = shLanc.getRange(first, COL.LIQUIDACAO, n, 1).getValues();

    let changes = 0;
    const outStatus = curStatusCol.map(r => [r[0]]);
    const outLiq    = curLiqCol.map(r => [r[0]]);

    for (let i = 0; i < n; i++) {
      const rowVals = bloc[i];
      const temAlgo = rowVals.some(v => v !== '' && v != null);
      if (!temAlgo || _estaFechadoPorComp_(rowVals)) continue;

      const forma  = rowVals[COL.FORMA - COL.DATA];
      const isCard = (typeof _findCartaoRow_ === 'function') ? !!_findCartaoRow_(forma)
                     : String(forma||'').toLowerCase().includes('cart');
      if (!isCard) continue;

      const d0 = (rowVals[0] instanceof Date && !isNaN(rowVals[0])) ? rowVals[0] : parseDateBR_(rowVals[0]);
      if (!d0) continue;

      let q = parseInt(rowVals[COL.PARCELAS - COL.DATA], 10); if (!isFinite(q) || q <= 0) q = 1;

      const cartRow = _findCartaoRow_(forma);
      if (!cartRow) continue;

      const nome = String(cartRow[0] || '').trim();
      const venc = _sanitizaDia_(cartRow[1]), ini = _sanitizaDia_(cartRow[2]), fim = _sanitizaDia_(cartRow[3]);
      if (venc == null || ini == null || fim == null) continue;

      // Todas as parcelas caem em faturas “pagas”?
      let allPaid = true;
      for (let p = 0; p < q; p++) {
        const dParc = new Date(d0); dParc.setMonth(dParc.getMonth() + p);
        const { mes, ano } = calcularCicloCartao(dParc, ini, fim, venc);
        const mesKey = Utilities.formatDate(new Date(ano, mes - 1, 1), tz, 'MM/yyyy');
        const fKey   = _cardKeyFromName_(nome) + '||' + mesKey;
        if (!paidKeys.has(fKey)) { allPaid = false; break; }
      }

      const curStatus = String(outStatus[i][0] || '').trim();
      const curLiq    = outLiq[i][0];

      if (allPaid) {
        if (!_isConciliadoStatus_(curStatus)) { outStatus[i][0] = 'Conciliado'; changes++; }
        if (!curLiq) { outLiq[i][0] = _today_(); changes++; }
      } else if (downgrade && _isConciliadoStatus_(curStatus)) {
        outStatus[i][0] = 'Pendente';
        outLiq[i][0] = '';
        changes++;
      }
    }

    if (changes) {
      shLanc.getRange(first, COL.STATUS, n, 1).setValues(outStatus);
      shLanc.getRange(first, COL.LIQUIDACAO, n, 1).setValues(outLiq);
      _maybeToast_(`🔁 Lançamentos sincronizados com o Resumo (${changes} ajuste${changes>1?'s':''}).`);
    }
    _logInfo_('sincronizarStatusLancamentosComResumo_', 'concluído', {changes});
  });
}

/* ===================== Reset simples (UI) ===================== */
function resetLancamentosBasico_(escopo){ // 'selecionadas' | 'tudo'
  const sh=_getLanc_(); if(!sh) return;
  const first=4, last=sh.getLastRow();
  const nCols=(COL_FP - COL.DATA + 1);
  if(last<first) return;

  let r0=first, r1=last;
  if(escopo==='selecionadas'){
    const rg=sh.getActiveRange();
    if(!rg){ _maybeToast_('Selecione as linhas a limpar.'); return; }
    r0=Math.max(first, rg.getRow());
    r1=Math.min(last, r0 + rg.getNumRows() - 1);
    if(r1<r0){ _maybeToast_('Nada a limpar.'); return; }
  }

  const num=r1-r0+1;
  const vals=sh.getRange(r0, COL.DATA, num, nCols).getValues();

  const blocks=[];
  let start=null, limpas=0, puladas=0;
  for(let i=0;i<num;i++){
    const podeLimpar = !_estaFechadoPorComp_(vals[i]);
    if(podeLimpar){
      if(start===null) start=r0+i;
    }else{
      puladas++;
      if(start!==null){ blocks.push([start, r0+i-1]); start=null; }
    }
  }
  if(start!==null) blocks.push([start, r0+num-1]);

  for(const [a,b] of blocks){
    const len=b-a+1;
    sh.getRange(a, COL.DATA,         len, nCols).clearContent().clearNote().setBackground(null);
    sh.getRange(a, COL.DETALHAMENTO, len, 1).clearDataValidations();
    limpas += len;
  }

  try{
    ensureDVStatusLancamento_();
    ensureDVNumericasLancamento_();
    if (typeof _ensureFormatsLanc_==='function') _ensureFormatsLanc_();
  }catch(_){}

  pendMarcarFull_(); pendMarcarPrevisao_(); pendMarcarFaturasResumo_();
  _maybeToast_(`🧹 Lançamentos limpos. Linhas limpas: ${limpas}${puladas?` | Puladas (mês fechado): ${puladas}`:''}`);
}

function openResetLancamentosUI_(){
  const ui=SpreadsheetApp.getUi();
  const btn=ui.alert('Reset de lançamentos','Limpar apenas as linhas selecionadas? (Não = limpar tudo)', ui.ButtonSet.YES_NO_CANCEL);
  if(btn===ui.Button.YES){
    resetLancamentosBasico_('selecionadas');
  } else if(btn===ui.Button.NO){
    const conf=ui.alert('Confirmar','Isto limpará TODOS os lançamentos (respeitando meses fechados). Continuar?', ui.ButtonSet.OK_CANCEL);
    if(conf===ui.Button.OK) resetLancamentosBasico_('tudo');
  }
}

/* ===================== MENUS & TRIGGERS (onOpen) ===================== */
/** ===================== MENU DINÂMICO (com Resets completos) ===================== **/

function onOpen(e){
  const ui = SpreadsheetApp.getUi();
  const G  = (typeof globalThis !== 'undefined') ? globalThis : this; // escopo global
  const menu = ui.createMenu('Financeiro');
  const faltando = [];

  const add = (label, fnName) => {
    if (typeof G[fnName] === 'function') { menu.addItem(label, fnName); return true; }
    faltando.push(`${label} → ${fnName}`); return false;
  };
  const addSep = () => { try{ menu.addSeparator(); }catch(_){} };
  const addSub = (titulo, itens) => {
    const sub = ui.createMenu(titulo); let ok=false;
    for (const [label, fn] of itens){
      if (typeof G[fn] === 'function'){ sub.addItem(label, fn); ok=true; }
      else faltando.push(`${titulo} › ${label} → ${fn}`);
    }
    if (ok) menu.addSubMenu(sub);
  };

  // ========= Ações rápidas
  add('⚡ Recalcular agora', 'doRecalcsNow_');
  addSep();

  // ========= Faturas & Resumo
  addSub('Faturas & Resumo', [
    ['Gerar/Atualizar Faturas', 'gerarFaturasDeCartao_'],
    ['Atualizar Resumo de Faturas', 'atualizarResumoFaturas_'],
    ['Rebuild "Parcelas do Cartão"', 'rebuildParcelasCartao_'],
    ['Sincronizar Status com Resumo', 'sincronizarStatusLancamentosComResumo_'],
    ['Validar Cartões (ciclo/limite)', 'validarCartoes_']
  ]);

  // ========= Previsão & Orçamento
  addSub('Previsão & Orçamento', [
    ['Atualizar Previsão (A:G)', 'atualizarPrevisaoCompleta_'],
    ['Projeções – próximos 3', 'preverProximos3_'],
    ['Projeções – próximos 6', 'preverProximos6_'],
    ['Projeções – próximos 12', 'preverProximos12_']
  ]);

  // ========= Auditorias
  addSub('Auditorias', [
    ['Duplicados (ID/Tx)', 'auditarDuplicados_'],
    ['Faturas × Resumo (parcial)', '_alertarErrosResumo_'],
    ['Auditoria Geral (Parte 2)', 'auditarGeralParte2_']
  ]);

  // ========= Resets & Manutenção (⭐ atualizado)
  addSub('Resets & Manutenção', [
    ['Reset — Lançamentos (competências/status/IDs/DV)', 'resetLancamentos_'],
    ['Reset — Faturas & Resumo (headers/gerar/rebuild/sync)', 'resetFaturasEResumo_'],
    ['Reset — DV (Categorias/Sub/Forma)', 'resetDV_'],
    ['Reset — Visual da Previsão (sparklines)', 'ensurePrevisaoProgressoVisual_'],
    ['Reset — Cond. Format (Previsão)', 'ensureCondFormatPrevisao_'],
    ['Reset — Cond. Format (Resumo)', 'ensureCondFormatResumoUtil_'],
    ['Reset — Proteções (meses fechados)', 'aplicarProtecaoLinhasFechadas_'],
    ['Reset — Normalizar Mês (Resumo B)', 'resetNormalizarResumo_'],
    ['Reset — Caches de Config', '_invalidateCfgCaches_'],
    ['RESET GERAL (seguro)', 'resetarTudo_']
  ]);

  // ========= Utilitários
  addSub('Utilitários', [
    ['Corrigir DV da Forma', 'corrigirDVFormaAgora_'],
    ['Fechar Mês Atual', 'fecharMesAtual_'],
    ['Reabrir Mês Atual', 'reabrirMesAtual_'],
    ['Gerar 300 lançamentos demo', 'gerar300LancamentosDemo_']
  ]);

  menu.addToUi();

  if (faltando.length){
    const msg = 'Itens não adicionados (função não encontrada):\n- ' +
                faltando.slice(0,15).join('\n- ') + (faltando.length>15 ? `\n… +${faltando.length-15}` : '');
    try { SS_().toast(msg); } catch(_){ try{ ui.alert('Financeiro — Itens indisponíveis', msg, ui.ButtonSet.OK); }catch(__){} }
  }
}
function onInstall(e){ onOpen(e); }

/** ===================== Resets específicos ===================== **/

// Reset — Lançamentos: recalcula competências (Caixa/Consumo), corrige status/liquidação,
// garante IDs e fingerprint, defaults de Centro/Tipo e revalida DVs dinâmicas.
function resetLancamentos_(){
  const sh=_getLanc_(), cfg=_getCfg_(); if(!sh||!cfg){ _maybeToast_('❌ Não achei a aba de Lançamentos/Config.'); return; }
  const first=4, last=sh.getLastRow(); if(last<first){ _maybeToast_('ℹ️ Sem linhas para reset.'); return; }

  const width = COL.ID_EXTRATO - COL.DATA + 1;
  const bloco = sh.getRange(first, COL.DATA, last-first+1, width).getValues();
  const centros = sh.getRange(first, COL.CENTRO_CUSTO, last-first+1, 1).getValues().flat();
  const tipos   = sh.getRange(first, COL.TIPO,       last-first+1, 1).getValues().flat();

  const idx=(c)=>c-COL.DATA;
  const hoje=_today_();

  const cartRowCache=new Map();
  const getCartRow=(forma)=>{
    const key=_normLower_(String(forma||'')); if(!key) return null;
    if(cartRowCache.has(key)) return cartRowCache.get(key);
    let r=null; try{ r=_findCartaoRow_(forma); }catch(_){}
    cartRowCache.set(key,r); return r;
  };

  const ccDef=_centroCustoDefault_(cfg);
  const tpDef=_tiposDefault_(cfg);

  // buffers
  const fmtData=[], compCx=[], fmtCx=[], noteCx=[], compCo=[], fmtCo=[];
  const sts=[], liq=[], ids=[], ccBuf=[], tpBuf=[], fps=[];

  for(let i=0;i<bloco.length;i++){
    const r=first+i, row=bloco[i]; if(!row.some(v=>v!==''&&v!=null)) continue;

    const dtRaw=row[idx(COL.DATA)];
    const d=(dtRaw instanceof Date && !isNaN(dtRaw)) ? _dateOnly_(dtRaw) : parseDateBR_(dtRaw);
    const forma=String(row[idx(COL.FORMA)]||'').trim();

    if(d){
      // consumo
      const compC=new Date(d.getFullYear(), d.getMonth(), 1);
      compCo.push({r,v:compC}); fmtCo.push({r,fmt:'MM/yyyy'});
      // caixa
      let compCaixa=compC, note='';
      const crow=getCartRow(forma);
      if(crow){
        const venc=_sanitizaDia_(crow[1]), ini=_sanitizaDia_(crow[2]), fim=_sanitizaDia_(crow[3]);
        if(venc!=null&&ini!=null&&fim!=null){
          const {mes,ano}=calcularCicloCartao(d, ini, fim, venc);
          compCaixa=new Date(ano, mes-1, 1);
        }else note='Ciclo inválido → mês da compra';
      }
      compCx.push({r,v:compCaixa}); fmtCx.push({r,fmt:'MM/yyyy'}); noteCx.push({r,note:note});
      fmtData.push({r,fmt:'dd/MM/yyyy'});
    }

    // status / liquidação coerentes
    let status=String(row[idx(COL.STATUS)]||'').trim();
    let liqVal=row[idx(COL.LIQUIDACAO)];
    const passado = d && _dateOnly_(d).getTime() <= _dateOnly_(hoje).getTime();
    if (forma && ehCartao_(forma)){
      if(!status){ status='Pendente'; sts.push({r,v:status}); }
      if(_isConciliadoStatus_(status) && !liqVal){ liqVal=_today_(); liq.push({r,v:liqVal}); }
    } else if (d){
      if (passado){
        if(!status){ status='Conciliado'; sts.push({r,v:status}); }
        if(_isConciliadoStatus_(status) && !liqVal){ liqVal=_today_(); liq.push({r,v:liqVal}); }
      } else {
        if(!_isConciliadoStatus_(status)){ status='Pendente'; sts.push({r,v:status}); }
        if(liqVal){ liq.push({r,v:''}); }
      }
    }

    // defaults Centro/Tipo
    if(!String(centros[i]||'').trim() && ccDef) ccBuf.push({r,v:ccDef});
    if(!String(tipos[i]||'').trim()   && tpDef) tpBuf.push({r,v:tpDef});

    // ID e fingerprint
    let id=String(row[idx(COL.ID_EXTRATO)]||'').trim();
    if(!id){
      id=_gerarIdExtrato_(sh, r, row, {returnOnly:true});
      ids.push({r,v:id});
      row[idx(COL.ID_EXTRATO)]=id;
    }
    fps.push({r,v:_rowFingerprint_(row)});
  }

  // escreve em bloco
  _setColFormatBatch_( sh, COL.DATA, fmtData);
  _setColValuesBatch_( sh, COL.COMPETENCIA, compCx);
  _setColFormatBatch_( sh, COL.COMPETENCIA, fmtCx);
  _setColNotesBatch_(  sh, COL.COMPETENCIA, noteCx);
  _setColValuesBatch_( sh, COL2.COMP_CONSUMO, compCo);
  _setColFormatBatch_( sh, COL2.COMP_CONSUMO, fmtCo);
  _setColValuesBatch_( sh, COL.STATUS, sts);
  _setColValuesBatch_( sh, COL.LIQUIDACAO, liq);
  _setColValuesBatch_( sh, COL.CENTRO_CUSTO, ccBuf);
  _setColValuesBatch_( sh, COL.TIPO, tpBuf);
  _setColValuesBatch_( sh, COL.ID_EXTRATO, ids);
  _setColValuesBatch_( sh, COL_FP, fps);

  try{ atualizarMenusDinamicos(sh, _getCfg_(), first, last-first+1); }catch(_){}
  _maybeToast_('♻️ Lançamentos resetados (competências, status, IDs, DV).');
}

// Reset — Faturas & Resumo: garante headers/formatos, gera faturas, atualiza resumo,
// normaliza mês, rebuild Parcelas e sincroniza status (se existir), reaplica CF.
function resetFaturasEResumo_(){
  try{ _ensureFaturasHeaders_(); }catch(_){}
  try{ _ensureResumoHeaders_(); }catch(_){}

  try{ gerarFaturasDeCartao_(); }catch(_){}
  try{ atualizarResumoFaturas_(); }catch(_){}
  try{
    const shRes=_safeEnsureSheet_(ABAS.RESUMO_FATURAS, ['Resumo de Faturas']);
    normalizeResumoMesColB_(shRes);
  }catch(_){}

  try{ if(typeof rebuildParcelasCartao_==='function') rebuildParcelasCartao_(); }catch(_){}
  try{ if(typeof sincronizarStatusLancamentosComResumo_==='function') sincronizarStatusLancamentosComResumo_({ downgrade:true }); }catch(_){}

  try{ ensureCondFormatResumoUtil_(); }catch(_){}
  _maybeToast_('♻️ Faturas & Resumo resetados.');
}

// ===== Aliases legados (compat com menus antigos)
function atualizarFaturasEResumo_(){ resetFaturasEResumo_(); }
function resetarLancamentos_(){ resetLancamentos_(); } // caso seu menu antigo use este nome

/** ======= Helpers de reset (idempotentes e seguros) ======= **/

// Recria as Data Validations de Lançamentos (Categorias/Sub/Forma)
function resetDV_(){
  const sh=_getLanc_(), cfg=_getCfg_(); if(!sh||!cfg) return;
  try{
    const cats=listaCategorias_(cfg);
    const subs=listaSubcategorias_(cfg);
    atualizarMenuCategoriasLancamento(sh, cats);
    atualizarMenuSubcategoriasLancamento(sh, subs);
    atualizarMenuFormasLancamento(sh, listaFormasComCartoes_(cfg));
    _maybeToast_('✔️ DV resetadas: Categoria/Sub/Forma.');
  }catch(e){ _maybeToast_('⚠️ resetDV_: '+(e&&e.message||e)); }
}

// Normaliza a coluna B (Mês) do Resumo para Date antes de ordenar
function resetNormalizarResumo_(){
  try{
    const shRes=_ensureResumoHeaders_();
    normalizeResumoMesColB_(shRes);
    _maybeToast_('✔️ Resumo: coluna Mês normalizada.');
  }catch(e){ _maybeToast_('⚠️ resetNormalizarResumo_: '+(e&&e.message||e)); }
}

// Reset geral (seguro): caches, DVs, visuais, CFs e proteções
function resetarTudo_(){
  try{ if(typeof _invalidateCfgCaches_==='function') _invalidateCfgCaches_(); }catch(_){}
  try{ resetDV_(); }catch(_){}
  try{ if(typeof ensurePrevisaoProgressoVisual_==='function') ensurePrevisaoProgressoVisual_(); }catch(_){}
  try{ if(typeof ensureCondFormatPrevisao_==='function') ensureCondFormatPrevisao_(); }catch(_){}
  try{ if(typeof ensureCondFormatResumoUtil_==='function') ensureCondFormatResumoUtil_(); }catch(_){}
  try{ if(typeof aplicarProtecaoLinhasFechadas_==='function') aplicarProtecaoLinhasFechadas_(); }catch(_){}
  try{ resetNormalizarResumo_(); }catch(_){}
  _maybeToast_('♻️ Reset geral concluído.');
}

/** (Opcional) Diagnóstico rápido do menu */
function debug_ListarFuncoesDeMenu_(){
  const ui = SpreadsheetApp.getUi();
  const G  = (typeof globalThis !== 'undefined') ? globalThis : this;
  const grupos = {
    'Ações Rápidas': ['doRecalcsNow_'],
    'Faturas & Resumo': ['gerarFaturasDeCartao_','atualizarResumoFaturas_','rebuildParcelasCartao_','sincronizarStatusLancamentosComResumo_','validarCartoes_'],
    'Previsão & Orçamento': ['atualizarPrevisaoCompleta_','preverProximos3_','preverProximos6_','preverProximos12_'],
    'Auditorias': ['auditarDuplicados_','_alertarErrosResumo_','auditarGeralParte2_'],
    'Resets & Manutenção': ['resetDV_','ensurePrevisaoProgressoVisual_','ensureCondFormatPrevisao_','ensureCondFormatResumoUtil_','aplicarProtecaoLinhasFechadas_','resetNormalizarResumo_','_invalidateCfgCaches_','resetarTudo_'],
    'Utilitários': ['corrigirDVFormaAgora_','fecharMesAtual_','reabrirMesAtual_','gerar300LancamentosDemo_']
  };
  const falt=[];
  Object.entries(grupos).forEach(([g, arr])=>arr.forEach(fn=>{ if (typeof G[fn] !== 'function') falt.push(`${g} › ${fn}`); }));
  const msg = falt.length ? ('Faltando:\n- ' + falt.join('\n- ')) : 'Tudo ok: todas as funções do menu estão disponíveis.';
  ui.alert('Diagnóstico do Menu', msg, ui.ButtonSet.OK);
}
/* Wrappers de menu */
function atualizarPrevisaoCompletaOverwrite_(){
  atualizarPrevisaoCompleta_({ overwriteMetas: true });
}
function verificarFaturasEResumo_(){
  _ensureFaturasHeaders_(); _ensureResumoHeaders_(); _maybeToast_('✅ Estruturas OK.');
}

/* ALTERADO: não chama mais Indicadores aqui; com alerta no fim */
function atualizarFaturasEResumo_(){
  gerarFaturasDeCartao_();
  atualizarResumoFaturas_();
  rebuildParcelasCartao_();
  sincronizarStatusLancamentosComResumo_({ downgrade: true });
  try{ alertasUtilizacaoEExposicao_(); }catch(_){}
  _maybeToast_('✅ Faturas & Resumo (e Parcelas) atualizados.');
}

function limparFaturasEResumo_(){
  const shFat=_ensureFaturasHeaders_();
  const shRes=_ensureResumoHeaders_();

  // Limpa conteúdos mantendo cabeçalho/formatações
  try{
    const lastF=shFat.getLastRow();
    if(lastF>1) shFat.getRange(2,1,lastF-1,11).clearContent();
  }catch(_){}
  try{
    const lastR=shRes.getLastRow();
    if(lastR>1) shRes.getRange(2,1,lastR-1,9).clearContent();
  }catch(_){}

  // Reprotege cabeçalho
  try{ _protectHeaderRow_(shFat); _protectHeaderRow_(shRes); }catch(_){}

  // Mantém a aba "Parcelas do Cartão" consistente
  try{ rebuildParcelasCartao_(); }catch(_){}

  // Limpa "Projeções" (se existir)
  try{
    const ss = SpreadsheetApp.getActive();
    const shP = ss.getSheetByName('Projeções') || ss.getSheetByName('Projecoes');
    if (shP) shP.clear();
  }catch(_){}

  _maybeToast_('🧹 Faturas/Resumo limpos (Parcelas atualizada e Projeções limpas).');
}

/* ===================== RECALCULAR (separado dos KPIs e de Faturas/Resumo) ===================== */
function recalcularTudo_(rapido, meses){
  if (rapido && Array.isArray(meses) && meses.length){
    pendMarcarMeses_(meses);         // só meses (mensais)
    // NÃO marca Previsão e NÃO marca Faturas/Resumo
  } else {
    pendMarcarFull_();               // full dos meses
    // NÃO marca Previsão e NÃO marca Faturas/Resumo
  }
  try { processarPendencias_(); }
  catch(e){ _maybeToast_('⚠️ Erro ao recalcular: ' + (e && e.message ? e.message : e)); }
}
function recalcularTudoFull_(){ try { recalcularTudo_(false); } catch(e){ _maybeToast_('⚠️ Erro ao recalcular (full): ' + (e && e.message ? e.message : e)); } }
function recalcularMesAtual_(){ try { const mes=(new Date()).getMonth()+1; recalcularTudo_(true, [mes]); } catch(e){ _maybeToast_('⚠️ Erro ao recalcular mês atual: ' + (e && e.message ? e.message : e)); } }

/** Corrige DV "Forma" em Lançamentos (fixas + cartões válidos) */
function corrigirDVFormaAgora_(){
  const shLanc = _getLanc_(), cfg = _getCfg_();
  if (!shLanc || !cfg) return;
  const formas = listaFormasComCartoes_(cfg); // remove duplicatas e exige ciclo válido

  const firstRow = 4;
  const lastUsed = Math.max(firstRow, shLanc.getLastRow());
  const buffer   = 400;

  const LIM = Math.max(1, Number(typeof LIMITE_LINHAS!=='undefined' ? LIMITE_LINHAS : 2000) || 2000);
  const totalRows = Math.min(LIM, (lastUsed - firstRow + 1) + buffer);
  if (totalRows<=0) return;

  const rng = shLanc.getRange(firstRow, COL.FORMA, totalRows, 1);
  if (!formas.length){ rng.clearDataValidations(); return; }
  _applyDVIfChanged_(rng, formas, false);
  _maybeToast_('✔ DV de "Forma" atualizada.');
}

/** Executa fn no máximo 1x a cada "intervalMs" por chave. */
function _runOnceEvery_(key, intervalMs, fn){
  try{
    const dp   = PropertiesService.getDocumentProperties();
    const prop = 'once_' + String(key||'').trim();
    const last = Number(dp.getProperty(prop) || 0);
    const now  = Date.now();
    const gap  = Math.max(60*1000, Number(intervalMs)||0); // mínimo 1 min
    if (now - last >= gap){
      if (typeof fn === 'function') { try{ fn(); }catch(_){ } }
      dp.setProperty(prop, String(now));
    }
  }catch(e){
    if (typeof fn === 'function') { try{ fn(); }catch(_){ } }
  }
}

/** Usa Projeções -> preenche Meta_Mensal/Meta Mês (match por Sub/Det). */
function metasUsarProjecaoSmart_(opts){
  opts = opts || {};
  const overwrite = !!opts.overwrite;

  const shPrev = getSheetSmart_(ABAS.PREVISAO_GASTOS, ['Previsao de Gastos']);
  const shProj = SpreadsheetApp.getActive().getSheetByName('Projeções') || SpreadsheetApp.getActive().getSheetByName('Projecoes');
  const shCfg  = _getCfg_();
  if(!shPrev || !shProj || !shCfg){ _maybeToast_('❌ Precisa das abas Previsão, Projeções e Configurações.'); return; }

  const norm = s => _normLower_(_stripDiacritics_(String(s||'')));

  function hdrMap(sh){
    const lastC = Math.max(1, sh.getLastColumn());
    const hdr = sh.getRange(1,1,1,lastC).getValues()[0].map(v=>String(v||'').trim());
    const map = new Map();
    hdr.forEach((h,i)=> map.set(norm(h), i+1));
    return { map, raw: hdr };
  }

  // mapa Item/Detalhamento -> Subcategoria (Config)
  const lastCfg = _cfgLastRow_(shCfg);
  const nRowsCfg = Math.max(0, lastCfg-2);
  const dets = nRowsCfg ? shCfg.getRange(3,5,nRowsCfg,1).getValues().flat() : []; // E
  const subsForDet = nRowsCfg ? shCfg.getRange(3,6,nRowsCfg,1).getValues().flat() : []; // F
  const subs = nRowsCfg ? shCfg.getRange(3,2,nRowsCfg,1).getValues().flat() : []; // B

  const itemToSub = new Map();
  const setMap = (key, sub) => {
    const k=norm(key), s=String(sub||'').trim();
    if(k && s) itemToSub.set(k, s);
  };
  (subs||[]).forEach(s => setMap(s, s));        // Sub -> Sub
  for (let i=0;i<dets.length;i++){              // Detalhamento -> Sub
    const d=dets[i], s=subsForDet[i]||'';
    if (d && s) setMap(d, s);
  }

  // Projeções: localizar colunas
  const projH = hdrMap(shProj);
  const colProjSub =
    projH.map.get('subcategoria') ||
    projH.map.get('sub-categoria') ||
    projH.map.get('sub categoria');

  let colProjPrev = null;
  for (let [key, idx] of projH.map.entries()){
    if (/^previs(a|ã)o mensal/.test(key)) { colProjPrev = idx; break; }
  }
  if(!colProjSub || !colProjPrev){
    _maybeToast_('❌ Na aba Projeções, preciso de "Subcategoria" e "Previsão Mensal (...)" .');
    return;
  }

  const lastProj = Math.max(2, shProj.getLastRow());
  const projVals = (lastProj>=2)
    ? shProj.getRange(2, Math.min(colProjSub,colProjPrev), lastProj-1, Math.abs(colProjPrev-colProjSub)+1).getValues()
    : [];
  const subPrev = new Map(); // sub normalizada -> valor previsto
  projVals.forEach(row=>{
    const s = row[colProjSub - Math.min(colProjSub,colProjPrev)];
    const v = row[colProjPrev - Math.min(colProjSub,colProjPrev)];
    const k = norm(s);
    if(k) subPrev.set(k, Number(v)||0);
  });

  // Previsão: localizar colunas
  const prevH = hdrMap(shPrev);

  const colPrevItem =
    prevH.map.get('item') || prevH.map.get('categoria');

  const colPrevMeta =
    prevH.map.get('meta_mensal') || prevH.map.get('meta mensal') ||
    prevH.map.get('meta mes')    || prevH.map.get('meta mês') ||
    prevH.map.get('meta mes (r$)') || prevH.map.get('meta mês (r$)') ||
    prevH.map.get('meta mês r$') || prevH.map.get('meta mes r$');

  if(!colPrevItem || !colPrevMeta){
    _maybeToast_('❌ Na Previsão, preciso de "Categoria/Item" e "Meta Mês (R$)/Meta_Mensal".');
    return;
  }

  const colPrevLock =
    prevH.map.get('meta_lock?') || prevH.map.get('meta_lock') ||
    prevH.map.get('meta lock?') || prevH.map.get('lock') ||
    prevH.map.get('meta travada?') || prevH.map.get('travar meta?');

  const colPrevManualSub =
    prevH.map.get('subcategoria_alvo') || prevH.map.get('subcategoria alvo') ||
    prevH.map.get('subcategoria-alvo') || prevH.map.get('forçar subcategoria') ||
    prevH.map.get('forcar subcategoria');

  const lastPrev = Math.max(2, shPrev.getLastRow());
  const items = (lastPrev>=2) ? shPrev.getRange(2, colPrevItem, lastPrev-1, 1).getValues().flat() : [];
  const metas  = (lastPrev>=2) ? shPrev.getRange(2, colPrevMeta, lastPrev-1, 1).getValues().flat()  : [];
  const locks  = (colPrevLock && lastPrev>=2) ? shPrev.getRange(2, colPrevLock, lastPrev-1, 1).getValues().flat() : null;
  const manual = (colPrevManualSub && lastPrev>=2) ? shPrev.getRange(2, colPrevManualSub, lastPrev-1, 1).getValues().flat() : null;

  const aliasKeys = Array.from(itemToSub.keys());

  const out = [];
  let escritos=0, puladosLock=0, semMatch=0, mantidos=0, comManual=0;

  for (let i=0; i<items.length; i++){
    const metaAtual = Number(metas[i]||0);
    const isLocked  = locks ? String(locks[i]||'').trim().toUpperCase() : '';
    if (isLocked==='SIM' || isLocked==='LOCK'){ out.push([metas[i]]); puladosLock++; continue; }

    let alvoSub = manual ? String(manual[i]||'').trim() : '';
    if (alvoSub) { comManual++; }
    else {
      const key = norm(items[i]);
      // (1) mapa direto (Sub ou Detalhamento)
      alvoSub = itemToSub.get(key) || '';
      // (2) fallback "contém" (só se der 1 único alvo)
      if (!alvoSub && key){
        const hits = aliasKeys.filter(k => k.includes(key) || key.includes(k));
        const uniqSubs = Array.from(new Set(hits.map(h => itemToSub.get(h)).filter(Boolean)));
        if (uniqSubs.length === 1) alvoSub = uniqSubs[0];
      }
    }

    if (alvoSub){
      const v = subPrev.get(norm(alvoSub));
      if (typeof v === 'number'){
        if (overwrite || !metaAtual){
          out.push([ _r2(v) ]); escritos++;
        } else { out.push([ metaAtual ]); mantidos++; }
      } else {
        out.push([ metaAtual ]); semMatch++;
      }
    } else {
      out.push([ metaAtual ]); semMatch++;
    }
  }

  if (out.length){
    shPrev.getRange(2, colPrevMeta, out.length, 1).setValues(out);
    try{ shPrev.getRange(2, colPrevMeta, out.length, 1).setNumberFormat('R$ #,##0.00'); }catch(_){}
  }

  _maybeToast_(
    '✅ Metas (smart): ' + escritos + ' escrita(s) ' +
    (mantidos ? '| Mantidas: ' + mantidos + ' ' : '') +
    (puladosLock ? '| Travadas: ' + puladosLock + ' ' : '') +
    (comManual ? '| Forçadas: ' + comManual + ' ' : '') +
    (semMatch ? '| Sem match: ' + semMatch : '')
  );
}

/** UI para metas */
function metasUsarProjecaoSmartUI_(){
  const ui=SpreadsheetApp.getUi();
  const btn=ui.alert('Usar Projeções → Metas (smart)',
    'Preencher "Meta Mês (R$)" com base na "Previsão Mensal" da aba Projeções.\n' +
    '- Match por Subcategoria/Detalhamento; fallback por "contém".\n' +
    '- Respeita "Meta_Lock?" (SIM/LOCK) e "Subcategoria_Alvo" (se existir).\n\n' +
    'Quer sobrescrever metas já preenchidas?',
    ui.ButtonSet.YES_NO_CANCEL);
  if(btn===ui.Button.YES){ metasUsarProjecaoSmart_({ overwrite:true }); }
  else if(btn===ui.Button.NO){ metasUsarProjecaoSmart_({ overwrite:false }); }
}

/* === NOVAS FUNÇÕES (faltantes) === */
function ativarVisualOrcamento_() {
  try {
    ensurePrevisaoProgressoVisual_(); // barras/percentuais H:I
    ensureCondFormatPrevisao_({});    // CF idempotente
    _maybeToast_('✅ Visual de orçamento (H:I) ativado.');
  } catch (e) {
    _maybeToast_('⚠️ Falha ao ativar visual: ' + (e && e.message ? e.message : e));
  }
}

function alertasUtilizacaoEExposicao_() {
  const tz   = _tz_();
  const hoje = _today_();
  const mesAtualStr = Utilities.formatDate(new Date(hoje.getFullYear(), hoje.getMonth(), 1), tz, 'MM/yyyy');

  const shRes = _ensureResumoHeaders_(); if (!shRes) return;

  const thr = (typeof utilThresholds_ === 'function')
    ? utilThresholds_()
    : (typeof UTIL_THRESH !== 'undefined' ? UTIL_THRESH : { WARN: 0.30, ALERT: 0.40 });

  let limites = new Map();
  try { const cfg = _getCfg_(); if (cfg) limites = _limitePorCartao_(cfg); } catch(_) {}

  const last = shRes.getLastRow();
  const rows = (last >= 2) ? shRes.getRange(2, 1, last-1, 9).getValues() : [];

  const avisos = [];
  for (const r of rows) {
    const cartao = String(r[0] || '').trim();
    if (!cartao) continue;

    const mesVal = r[1];
    const mesStr = (mesVal instanceof Date && !isNaN(mesVal))
      ? Utilities.formatDate(new Date(mesVal.getFullYear(), mesVal.getMonth(), 1), tz, 'MM/yyyy')
      : String(mesVal || '').trim();
    if (mesStr !== mesAtualStr) continue;

    const liquido = Number(r[2] || 0);
    const util = (r[7] === '' || r[7] == null)
      ? (function(){ const lim = Number(limites.get(cartao) || 0); return lim > 0 ? (liquido / lim) : null; })()
      : Number(r[7]);
    if (util == null) continue;

    const lim = Number(limites.get(cartao) || 0);
    const utilPct = Math.round(util * 100);

    if (util >= thr.ALERT) {
      avisos.push(`🚨 ${cartao}: ${utilPct}% ${lim>0?`(R$ ${_r2(liquido).toFixed(2)} / R$ ${_r2(lim).toFixed(2)})`:''}`);
    } else if (util >= thr.WARN) {
      avisos.push(`⚠️ ${cartao}: ${utilPct}% ${lim>0?`(limite R$ ${_r2(lim).toFixed(2)})`:''}`);
    }
  }

  if (avisos.length) _maybeToast_(`Cartões acima do uso seguro:\n- ${avisos.join('\n- ')}`);
}
