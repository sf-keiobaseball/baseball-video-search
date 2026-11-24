// STEP 39: 野手・投手とも再生可能なアングルが1つ以上ある投球だけを検索結果に表示。候補とカレンダーも動画ありに限定。
// STEP 38: 入口のタイトル下の説明を削除。投手アイコンはユーザー提供の実写野球ボール写真（WebP内蔵）に変更。
// STEP 37: 入口の選択カードから説明文を削除し、野手=バット・投手=野球ボールのSVGイラストに変更。
// STEP 36: 入口で野手／投手を選択。野手7項目を維持、投手8項目・4アングル／2画面対応。
// 投手データは同じExcelの慶應投手側の行から作る。リーグ戦は各視点で別個に同期登録。
// STEP 35: 1画面・2画面の両方で各動画の直下に独立したアングル選択パネルを配置。
// STEP 34: 日付の複数選択カレンダーと、左右2画面での独立した動画選択・再生を追加。
// STEP 33: 全7項目をチェックボックスで複数選択可能に変更。同項目OR・異項目AND、候補連動。
// STEP 32: 打撃結果フィルターをストライク・ボール・ファウル・凡打・安打の5分類に変更。
// STEP 31: 5シート入りExcel (.xlsx) を1回で読み込み。CSVとの後方互換を維持。
// リーグ戦タイムスタンプ必須。公開済み同期点を照合して引き継ぐ。
// STEP 30: 試合ごとにIndexedDBへ一時保存。更新終了で一括クラウド公開。
// STEP 22: Supabaseクラウド保存・管理者認証。
// STEP 17: 閲覧者の案内欄を非表示・試合区分を追加・全条件の一括解除。
// STEP 16: 管理者・閲覧者のローカル分離と更新終了（IndexedDB）。外部公開・認証は未実装。
// STEP 15: 紅白戦CSV追加・ 連動型フィルター・6項目検索（対戦大学追加）・4区分・YouTube再生・リーグ戦チャプター同期（余裕3秒）
// CSVはローカル処理。動画を選択するとYouTubeに接続します。
'use strict';
const $ = id => document.getElementById(id);
const adminMode = new URLSearchParams(location.search).get('admin') === '1';
let dirty = false, busy = 0, publishing = false, ready = false;
let leagueDraftPending=()=>false;
let resumeMatch='';
const uploadStatuses = new Map();
function markDirty(){invalidateVideoSearchCache();dirty=true;updatePublishUI();}

const labels = {league:'リーグ戦', practice:'OP戦(A)', practiceB:'OP戦(B)', intraSquad:'紅白戦', sheet:'シートBT'};
const isPractice = category => category === 'practice' || category === 'practiceB' || category === 'intraSquad';
const angleLabel = key => key === 'センター裏' ? 'センター裏(Big6)' : key;
const typeMap = {Fastball:'ストレート', Slider:'スライダー', Cutter:'カットボール', Splitter:'フォーク', CurveBall:'カーブ', TwoSeamFastball:'ツーシーム', ChangeUp:'チェンジアップ', Sweeper:'スイーパー'};
const clean = v => String(v ?? '').trim();
// STEP 9: ユーザー指定の検索用分類。元の投球・打席結果は保持。
const resultGroups = {
  "ストライク": [
    "見送り",
    "空振り",
    "三振",
    "空三振",
    "見三振",
    "バ三振",
    "見送り三振",
    "見逃し三振",
    "振逃"
  ],
  "ボール": [
    "ボール",
    "四球",
    "死球"
  ],
  "ファウル": [
    "ファウル",
    "ファール"
  ],
  "凡打": [
    "一ゴ",
    "二ゴ",
    "三ゴ",
    "遊ゴ",
    "投ゴ",
    "捕ゴ",
    "一ゴロ",
    "二ゴロ",
    "三ゴロ",
    "遊ゴロ",
    "投ゴロ",
    "一飛",
    "二飛",
    "三飛",
    "遊飛",
    "投飛",
    "捕飛",
    "左飛",
    "中飛",
    "右飛",
    "一直",
    "二直",
    "三直",
    "遊直",
    "投直",
    "左直",
    "中直",
    "右直",
    "一邪直",
    "一邪飛",
    "二邪飛",
    "三邪飛",
    "遊邪飛",
    "捕邪飛",
    "左邪飛",
    "右邪飛",
    "邪飛",
    "一併",
    "二併",
    "三併",
    "遊併",
    "投併",
    "捕併",
    "遊併殺",
    "投ギ",
    "投犠",
    "投犠打",
    "一犠",
    "一犠打",
    "三犠打",
    "捕犠",
    "捕犠打",
    "犠打",
    "右犠",
    "中犠",
    "左犠飛",
    "中犠飛",
    "右犠飛",
    "一失",
    "二失",
    "三失",
    "遊失",
    "中失",
    "一失策",
    "二失策",
    "三失策",
    "遊失策",
    "投失策",
    "左失策",
    "右失策",
    "遊野選",
    "投ギ野",
    "捕ギ野"
  ],
  "安打": [
    "安打",
    "左安",
    "中安",
    "右安",
    "一安",
    "二安",
    "三安",
    "遊安",
    "投安",
    "捕安",
    "左２",
    "中２",
    "右２",
    "三２",
    "捕２",
    "左2",
    "中2",
    "右2",
    "２塁打",
    "左二安",
    "右二安",
    "左３",
    "中３",
    "右３",
    "中3",
    "中三安",
    "右三安",
    "左中三安",
    "左本",
    "中本",
    "右本",
    "本塁打"
  ]
};
const resultClassMap = new Map(Object.entries(resultGroups).flatMap(([group, values]) => values.map(value => [value.normalize('NFKC'), group])));
function classifyResult(value) {
    return resultClassMap.get(clean(value).normalize('NFKC')) || '';
}

function parseCSV(text) {
    text = text.replace(/^\uFEFF/, '');
    const out = []; let row = [], field = '', quoted = false;
    for (let i=0;i<text.length;i++) {
        const c=text[i];
        if (quoted) {
            if(c==='"') { if(text[i+1]==='"') {field+='"';i++;} else quoted=false; }
            else field+=c;
        } else if(c==='"' && field==='') quoted=true;
        else if(c===',') {row.push(field);field='';}
        else if(c==='\n'||c==='\r') {
            if(c==='\r' && text[i+1]==='\n') i++;
            row.push(field);out.push(row);row=[];field='';
        } else field+=c;
    }
    if(quoted) throw new Error('CSVの引用符が閉じていません。');
    if(row.length || field!=='') {row.push(field);out.push(row);}
    return out;
}
function normalizeCSV(text, category) {
    return normalizeRows(parseCSV(text),category);
}
function normalizeRows(table, category) {
    const rows=table.slice(), headers=(rows.shift()||[]).map(clean);
    const sheet=category==='sheet';
    const required=sheet?['日付','打者','投手','Pitch Type','Velocity']:['日付','打者名','投手名','打者チーム','投手チーム','球種','球速'];
    if(category==='league')required.push('タイムスタンプ');
    const missing=required.filter(k=>!headers.includes(k));
    if(missing.length) throw new Error('CSVの区分または形式が違います。不足列：'+missing.join('、'));
    if(isPractice(category) && !headers.includes('ネット裏')) throw new Error('OP戦・紅白戦用のCSVを選択してください（ネット裏列が必要）。');
    if(category==='league' && headers.includes('ネット裏')) throw new Error('リーグ戦用のCSVを選択してください。');
    const indexes=new Map();headers.forEach((h,i)=>{if(!indexes.has(h)) indexes.set(h,i);});
    const normalized=[];
    for(const [sourceRow,row] of rows.entries()) {
        const get=k=>clean(row[indexes.get(k)]);
        const batter=get(sheet?'打者':'打者名');
        const dateMatch=get('日付').match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
        if(!batter||!dateMatch) continue;
        // 同じExcelに慶應の攻撃と守備の両方が含まれる。シートBTには所属列がないため両画面で利用可能。
        const roles=sheet?['batter','pitcher']:[...(get('打者チーム')==='慶應'?['batter']:[]),...(get('投手チーム')==='慶應'?['pitcher']:[])];
        if(!roles.length)continue;
        const date=dateMatch[1]+'/'+dateMatch[2].padStart(2,'0')+'/'+dateMatch[3].padStart(2,'0');
        const side=v=>v==='右'?'R':v==='左'?'L':'';
        const rawType=get(sheet?'Pitch Type':'球種');
        const rawSpeed=Number(get(sheet?'Velocity':'球速'));
        const measurementValid=!sheet||get('投手誤計測').toUpperCase()!=='TRUE';
        const speed=measurementValid&&Number.isFinite(rawSpeed)&&rawSpeed>0?rawSpeed:null;
        const pitchResult=get('投球結果')==='見逃し'?'見送り':get('投球結果')==='ファール'?'ファウル':get('投球結果');
        const atBatResult=get('打席結果');
        const balls=get('ボールカウント'), strikes=get('ストライクカウント');
        const angles={};
        if(isPractice(category)) for(const name of ['ネット裏','センター裏','一塁側','三塁側','一塁側(投手)','三塁側(投手)']) {const url=get(name);if(url) angles[name]={url};}
        if(sheet) for(const [name,prefix] of Object.entries({'ネット裏':'BackNet','センター裏':'Center','一塁側':'First','三塁側':'Third','一塁側(投手)':'FirstPitcher','三塁側(投手)':'ThirdPitcher'})) {const url=get(prefix+'Link');if(url) angles[name]={url,seconds:get(prefix+'Second')};}
        for(const role of roles) normalized.push({role,sourceRow,pitchInAtBat:get('打席内球数'),category, date, year:dateMatch[1], opponent:sheet?'':role==='pitcher'?get('打者チーム'):get('投手チーム'), batter, pitcher:get(sheet?'投手':'投手名'), batterSide:sheet?'':side(get('打者打席左右')), pitcherSide:sheet?'':side(get('投手投げ手左右')), inning:sheet||!get('イニング')?'':get('イニング')+'回'+get('表/裏'), count:sheet||balls===''||strikes===''?'':balls+'-'+strikes, pitchType:measurementValid?(typeMap[rawType]||rawType):'', speed, pitchResult, atBatResult, result:atBatResult||pitchResult, resultClass:classifyResult(atBatResult||pitchResult), timestamp:sheet?get('Time'):get('タイムスタンプ'), angles});
    }
    return normalized;
}
// STEP 31: Excelは管理者のブラウザ内で解析する（Excel本体はSupabaseに送信しない）。
// SheetJS 0.20.3 の公式配布版を必要時だけ読み込む。既存の index.html への追記は不要。
const EXCEL_LIBRARY_URL='https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
let excelLibraryLoading=null;
function loadExcelLibrary(){
    if(window.XLSX)return Promise.resolve(window.XLSX);
    if(excelLibraryLoading)return excelLibraryLoading;
    excelLibraryLoading=new Promise((resolve,reject)=>{
        const script=document.createElement('script');script.src=EXCEL_LIBRARY_URL;
        script.async=true;
        script.onload=()=>window.XLSX?resolve(window.XLSX):reject(new Error('Excel解析ライブラリを読み込めませんでした。'));
        script.onerror=()=>reject(new Error('Excel解析ライブラリへの接続に失敗しました。ネット接続・通信制限を確認してください。'));
        document.head.append(script);
    }).catch(err=>{excelLibraryLoading=null;throw err;});
    return excelLibraryLoading;
}
const twoDigits=n=>String(n).padStart(2,'0');
function fromExcelSerial(value,kind='date'){
    // Excel 1900日付システム（通常の.xlsx）。日付/日時は常に日本の壁時計時刻として扱う。
    if(typeof value!=='number'||!Number.isFinite(value))return value;
    if(kind==='time'&&value>=0&&value<1){
        const sec=((Math.round(value*86400)%86400)+86400)%86400;
        return `${twoDigits(Math.floor(sec/3600))}:${twoDigits(Math.floor(sec/60)%60)}:${twoDigits(sec%60)}`;
    }
    if(value<20000||value>=80000)return value;
    const d=new Date(Date.UTC(1899,11,30)+Math.round(value*86400000));
    if(!Number.isFinite(d.valueOf()))return value;
    const date=`${d.getUTCFullYear()}/${twoDigits(d.getUTCMonth()+1)}/${twoDigits(d.getUTCDate())}`;
    return kind==='datetime'?date+' '+[d.getUTCHours(),d.getUTCMinutes(),d.getUTCSeconds()].map(twoDigits).join(':'):date;
}
function readExcelCategory(workbook, category, sheetName){
    const XLSX=window.XLSX, sheet=workbook.Sheets[sheetName];
    if(!sheet)throw new Error('Excelに「'+sheetName+'」シートがありません。シート名を確認してください。');
    const rows=XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:'',blankrows:false});
    if(rows.length<2)throw new Error('「'+sheetName+'」にデータがありません。');
    const header=rows[0].map(clean);
    const dateCol=header.indexOf('日付'),stampCol=header.indexOf('タイムスタンプ'),timeCol=header.indexOf('Time');
    if(category==='league'&&stampCol<0){
        throw new Error('リーグ戦の「タイムスタンプ」列がありません。提供した時刻修正版Excelを使用してください。');
    }
    // バックアップしやすいよう元の.xlsx/ワークシートは変更しない。
    for(let i=1;i<rows.length;i++){
        const r=rows[i];
        if(dateCol>=0)r[dateCol]=fromExcelSerial(r[dateCol],'date');
        if(stampCol>=0)r[stampCol]=fromExcelSerial(r[stampCol],'datetime');
        if(category==='sheet'&&timeCol>=0)r[timeCol]=fromExcelSerial(r[timeCol],'time');
    }
    const normalized=normalizeRows(rows,category);
    if(!normalized.length)throw new Error('「'+sheetName+'」は読み込み対象が0件です。日付・打者チーム等を確認してください。');
    if(category==='league'&&normalized.some(r=>clockSeconds(r.timestamp)===null)){
        throw new Error('リーグ戦に無効なタイムスタンプがあります。年月日と時分秒を確認してください。');
    }
    return normalized;
}
const dataByCategory={league:[],practice:[],practiceB:[],intraSquad:[],sheet:[]};
let results=[], page=1, selectedIndex=-1, selectedAngle=null;
let activeView='batter';
const recordView=r=>r.role||'batter';
const allowedAngles=view=>view==='pitcher'?['ネット裏','センター裏','一塁側(投手)','三塁側(投手)']:['ネット裏','センター裏','一塁側','三塁側','一塁側(投手)','三塁側(投手)'];
const leagueDraftChecks=[];
const leaguePanels=[];
const calendarState={month:'',hadDates:false};
const form=$('searchForm');
// 既存HTMLを変更せず検索フォームだけを再構築する。
const filterFields = [['year','年度'],['date','日付'],['category','試合区分'],['opponent','対戦大学'],['pitcher','投手'],['batter','打者'],['pitchType','球種'],['result','打撃結果']];
// STEP 33: 全7項目でチェックボックス式の複数選択を可能にする。
// 同じ項目はOR、異なる項目はAND。元の投球データや公開データの保存形式は変更しない。
// 保存済みデータの古い分類を参照せず、最新の5分類で都度判定する。
const filterValue = (record,key) => key==='result' ? classifyResult(record.result || record.atBatResult || record.pitchResult) : record[key];
const filterSelections=Object.fromEntries(filterFields.map(([key])=>[key,new Set()]));
const filterPanels=Object.create(null);
const filterStyle=document.createElement('style');
filterStyle.textContent=`
.bvs-multi-filter{margin-top:6px;background:#fff;border:1px solid #b9c5d3;border-radius:6px;color:#172b46;overflow:hidden}
.bvs-multi-filter>summary{display:block;list-style:none;cursor:pointer;position:relative;padding:10px 30px 10px 12px;min-height:42px;line-height:1.45;font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bvs-multi-filter>summary::-webkit-details-marker{display:none}
.bvs-multi-filter>summary::after{content:'▾';position:absolute;right:12px;top:10px;color:#24558e}
.bvs-multi-filter[open]>summary::after{content:'▴'}
.bvs-multi-filter[open]>summary{background:#eef4fb}
.bvs-multi-filter>summary:focus-visible{outline:2px solid #24558e;outline-offset:-2px}
.bvs-multi-body{padding:8px;border-top:1px solid #d8e1ec}
.bvs-multi-search{display:block;width:100%;min-width:0;box-sizing:border-box;padding:8px;border:1px solid #b9c5d3;border-radius:4px;font-size:13px}
.bvs-multi-options{max-height:220px;overflow-y:auto;overscroll-behavior:contain;margin-top:7px}
.bvs-multi-option{display:flex!important;align-items:center;gap:8px;padding:7px 4px;cursor:pointer;line-height:1.4;font-size:13px;font-weight:normal!important;overflow-wrap:anywhere}
.bvs-multi-option:hover{background:#f0f5fa}
.bvs-multi-option input{flex:0 0 auto;width:16px;height:16px;margin:0;accent-color:#24558e}
.bvs-multi-option.bvs-unavailable{color:#a33c32}
.bvs-multi-meta{font-size:11px;color:#64748b;margin:6px 0}
.bvs-multi-clear{font-size:12px;padding:5px 7px;border:1px solid #b9c5d3;border-radius:4px;background:white;color:#24558e;cursor:pointer}
.bvs-multi-clear:disabled{opacity:.45;cursor:default}
`;
// STEP 34: 日付は月カレンダーで複数選択。動画は1画面/2画面を切り替える。
filterStyle.textContent+=`
.bvs-calendar-host{max-height:none;overflow:visible;margin-top:4px}
.bvs-calendar-head{display:flex;align-items:center;gap:5px;margin:5px 0 12px}
.bvs-calendar-head button{flex:0 0 32px;min-height:33px;border:1px solid #bbc9db;background:#eef4fa;border-radius:5px;cursor:pointer}
.bvs-calendar-head button:disabled{opacity:.35;cursor:default}
.bvs-calendar-head input[type=month]{flex:1;min-width:0;max-width:100%;padding:7px 4px;border:1px solid #bbc9db;border-radius:5px;font-size:13px;background:#fff;color:#172b46}
.bvs-calendar-days{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:3px}
.bvs-calendar-week{font-size:11px;text-align:center;color:#64748b;font-weight:700;padding:5px 0}
.bvs-calendar-days button{min-width:0;min-height:32px;border:1px solid transparent;border-radius:5px;background:#edf5fe;color:#172b46;font-size:12px;cursor:pointer}
.bvs-calendar-days button:hover:not(:disabled){border-color:#24558e}
.bvs-calendar-days button.bvs-calendar-selected{background:#24558e;color:#fff;font-weight:bold;border-color:#24558e}
.bvs-calendar-days button.bvs-calendar-unavailable:not(.bvs-calendar-selected){background:#f5f5f5;color:#959ca7}
.bvs-calendar-days button:disabled{background:transparent;color:#c3cad1;cursor:default}
.bvs-calendar-days button:focus-visible{outline:2px solid #24558e;outline-offset:2px}
.bvs-calendar-hint{font-size:11px;line-height:1.5;color:#64748b;margin:8px 0}
.bvs-calendar-picked{display:flex;flex-wrap:wrap;gap:5px;max-height:95px;overflow:auto;margin:8px 0}
.bvs-calendar-picked button{border:1px solid #9fb8d4;border-radius:20px;padding:4px 8px;color:#24558e;background:#eef4fb;font-size:11px;cursor:pointer}
.bvs-mode-bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:8px 0 12px;color:#172b46}
.bvs-mode-bar strong{margin-right:4px;font-size:14px}
.bvs-mode-button{border:1px solid #b9c5d3;border-radius:6px;background:#fff;color:#172b46;padding:8px 12px;cursor:pointer;font-size:13px}
.bvs-mode-button.active{background:#24558e;border-color:#24558e;color:white}
.bvs-mode-explain{color:#63738a;font-size:12px;line-height:1.45}
.bvs-video-stage{display:grid;grid-template-columns:minmax(0,1fr);gap:14px;margin-bottom:10px}
.bvs-video-stage[data-count="2"]{grid-template-columns:repeat(2,minmax(0,1fr))}
.bvs-dual-slot{min-width:0}
.bvs-dual-title{font-weight:700;font-size:13px;color:#172b46;margin:5px 0 7px}
.bvs-video-stage[data-count="1"] .bvs-dual-title{display:none}
.bvs-video-stage[data-count="2"] #videoPlayer,.bvs-video-stage[data-count="2"] #bvsSecondPlayer{width:100%!important;max-width:100%;height:auto!important;min-height:0!important;aspect-ratio:16/9;box-sizing:border-box}
.bvs-video-stage iframe{display:block;width:100%!important;height:100%!important;border:0}
#bvsSecondPlayer{position:relative;aspect-ratio:16/9;background:#101b2e;overflow:hidden;border-radius:6px;color:white}
#bvsSecondPlayer .video-placeholder{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;box-sizing:border-box;text-align:center}
.bvs-second-detail{font-size:12px;color:#54657a;margin:8px 0;line-height:1.5;overflow-wrap:anywhere}
.bvs-second-actions{display:flex;align-items:center;gap:7px;flex-wrap:wrap;margin:9px 0}
.bvs-second-actions button,.bvs-compare-angle{border:1px solid #b9c5d3;border-radius:5px;background:#fff;color:#24558e;padding:7px 9px;cursor:pointer;font-size:12px}
.bvs-second-actions button:disabled,.bvs-compare-angle:disabled{opacity:.45;cursor:default}
.bvs-second-angle-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:5px;margin:7px 0}
.bvs-second-angle-list button.active{background:#24558e;color:white}
/* STEP 35: 両方の映像のすぐ下にアングルを表示する */
.bvs-dual-angle-title{font-size:12px;font-weight:700;color:#172b46;margin:10px 0 6px}
.bvs-dual-angle-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:6px;margin:0 0 12px}
.bvs-dual-angle-list button{display:block;box-sizing:border-box;width:100%;min-width:0;min-height:37px;padding:7px 5px;border:1px solid #b9c5d3;border-radius:5px;background:#fff;color:#24558e;font-size:12px;line-height:1.3;white-space:normal;cursor:pointer}
.bvs-dual-angle-list button:hover:not(:disabled){border-color:#24558e;background:#eef4fb}
.bvs-dual-angle-list button.active{background:#24558e!important;color:#fff!important;border-color:#24558e!important}
.bvs-dual-angle-list button:disabled{opacity:.42;cursor:default}
.bvs-dual-angle-list button:focus-visible{outline:2px solid #24558e;outline-offset:2px}
.bvs-video-stage[data-count="1"] .bvs-first-angle-list{grid-template-columns:repeat(3,minmax(0,1fr))}
@media(max-width:560px){.bvs-video-stage[data-count="1"] .bvs-first-angle-list{grid-template-columns:repeat(2,minmax(0,1fr))}}

.bvs-second-link{display:inline-block;margin:4px 0 7px;font-size:12px;color:#24558e;overflow-wrap:anywhere}
.bvs-row-actions{display:flex;flex-wrap:wrap;gap:4px;align-items:center}
.bvs-row-actions button{white-space:nowrap}
.bvs-selected-second{outline:2px solid #668ab6;outline-offset:-2px}
@media(max-width:900px){.bvs-video-stage[data-count="2"]{grid-template-columns:minmax(0,1fr)}}
`;
document.head.append(filterStyle);
form.replaceChildren();
for(const [key,name] of filterFields){
    const group=document.createElement('div');group.className='filter-group';
    const label=document.createElement('label');label.textContent=name;
    const details=document.createElement('details');details.id='multi-'+key;details.className='bvs-multi-filter';
    const summary=document.createElement('summary');summary.textContent='すべて';
    const body=document.createElement('div');body.className='bvs-multi-body';
    const query=document.createElement('input');query.type='search';query.className='bvs-multi-search';query.placeholder='候補を検索';query.setAttribute('aria-label',name+'の候補を検索');
    const meta=document.createElement('div');meta.className='bvs-multi-meta';
    const list=document.createElement('div');list.className=key==='date'?'bvs-calendar-host':'bvs-multi-options';
    if(key==='date'){query.style.display='none';query.setAttribute('aria-hidden','true');}
    const clear=document.createElement('button');clear.type='button';clear.className='bvs-multi-clear';clear.textContent='この項目の選択を解除';
    clear.addEventListener('click',()=>{filterSelections[key].clear();populateFilters();});
    query.addEventListener('input',()=>renderFilterOptions(key));
    body.append(query,meta,list,clear);details.append(summary,body);group.append(label,details);form.append(group);
    filterPanels[key]={group,details,summary,query,meta,list,clear,values:[],available:new Set()};
    details.addEventListener('toggle',()=>{
        if(details.open)for(const [otherKey,panel] of Object.entries(filterPanels))if(otherKey!==key)panel.details.open=false;
    });
}
const searchButton=document.createElement('button');searchButton.type='button';searchButton.id='searchButton';searchButton.className='search-button';searchButton.textContent='検索する';
const resetButton=document.createElement('button');resetButton.type='button';resetButton.id='clearFilters';resetButton.className='reset-button';resetButton.textContent='絞り込みを全解除';
form.append(searchButton,resetButton);
// テスト用CSV選択欄を追加。
const panel=document.createElement('details');
panel.open=true;
panel.style.cssText='margin-bottom:20px;padding:12px;background:#fff;border:1px solid #b9c5d3;border-radius:6px;font-size:12px;';
const summary=document.createElement('summary');summary.textContent='野手・投手動画データ読み込み（Excel 1ファイル対応）';summary.style.cssText='font-weight:bold;cursor:pointer;margin-bottom:8px;';panel.append(summary);
const excelHint=document.createElement('p');excelHint.textContent='5つのシート（リーグ戦／OP戦(A)／OP戦(B)／紅白戦／シートBT）が入ったExcelを1つ選択してください。慶應の打撃・投球データをどちらも読み取ってブラウザ内に一時保存し、「更新終了・一括保存」で初めて閲覧者側へ反映します。';excelHint.style.cssText='line-height:1.6;margin:8px 0 10px';panel.append(excelHint);
const excelLabel=document.createElement('label');excelLabel.textContent='5試合区分まとめて読み込み（.xlsx）';excelLabel.style.cssText='display:block;font-weight:bold;margin:10px 0';
const excelInput=document.createElement('input');excelInput.type='file';excelInput.accept='.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';excelInput.style.cssText='display:block;width:100%;font-size:12px;margin-top:8px';excelLabel.append(excelInput);panel.append(excelLabel);
const excelStatus=document.createElement('p');excelStatus.textContent='Excel未読み込み';excelStatus.setAttribute('aria-live','polite');excelStatus.style.cssText='white-space:pre-line;line-height:1.5;overflow-wrap:anywhere';panel.append(excelStatus);
const legacyCSV=document.createElement('details');legacyCSV.style.cssText='margin-top:14px;border-top:1px solid #cbd5e1;padding-top:10px';
const legacyTitle=document.createElement('summary');legacyTitle.textContent='従来のCSVを区分別に読み込む（必要な場合のみ）';legacyTitle.style.cssText='font-weight:bold;cursor:pointer;';legacyCSV.append(legacyTitle);panel.append(legacyCSV);
const hint=document.createElement('p');hint.textContent='CSVでは更新する区分だけ選択できます。他の区分は保持します。Excelでは全5区分をまとめて差し替えます。更新終了まで閲覧者ページには反映されません。';hint.style.marginBottom='12px';legacyCSV.append(hint);
for(const category of Object.keys(labels)) {
    const label=document.createElement('label');label.textContent=labels[category];label.style.cssText='display:block;font-weight:bold;margin-top:10px;';
    const input=document.createElement('input');input.type='file';input.accept='.csv,text/csv';input.style.cssText='display:block;width:100%;font-size:11px;margin:6px 0;';label.append(input);legacyCSV.append(label);
    const status=document.createElement('p');status.textContent='未読み込み';status.setAttribute('aria-live','polite');status.style.cssText='color:#64748b;overflow-wrap:anywhere;';legacyCSV.append(status);uploadStatuses.set(category,status);
    input.addEventListener('change',async()=>{
        const file=input.files[0];if(!file||!ready||publishing||busy){input.value='';return;}
        if(leagueDraftPending()){alert('先にリーグ戦の入力を一時保存してください。');input.value='';return;}
        const oldRows=dataByCategory[category],oldAnchors=new Map(leagueAnchors),oldDirty=dirty;
        status.textContent='読み込み中…';input.disabled=true;busy++;const leagueBox=panel.nextElementSibling;if(leagueBox)leagueBox.inert=true;updatePublishUI();
        try {
            const records=normalizeCSV(await file.text(),category);
            const remap=category==='league'?remapLeagueAnchors(dataByCategory.league,records):null;
            dataByCategory[category]=records;
            if(remap){leagueAnchors.clear();invalidateVideoSearchCache();for(const [k,v] of remap.entries)leagueAnchors.set(k,v);}
            await persistDraft();
            if(remap)refreshLeaguePanel();
            markDirty();
            status.textContent=`${records.length.toLocaleString('ja-JP')}件 読み込み完了`+(category==='sheet'?'（所属未確認）':'')+(remap?`｜リーグ戦基準点：${remap.kept}件引継ぎ、${remap.dropped}件要再登録`:'');
            status.style.color='#17633a';
            populateFilters();search();
        } catch(e) {dataByCategory[category]=oldRows;leagueAnchors.clear();invalidateVideoSearchCache();for(const [k,v] of oldAnchors)leagueAnchors.set(k,v);dirty=oldDirty;status.textContent='読み込み失敗：'+e.message+'（以前のデータは保持）';status.style.color='#b42318';}
        finally {input.disabled=false;input.value='';busy--;if(leagueBox)leagueBox.inert=false;updatePublishUI();}
    });
}
// 5シートがすべて正しく読めることを確認してから一括差し替え。失敗時は以前のデータを保持。
excelInput.addEventListener('change',async()=>{
    const file=excelInput.files?.[0];
    if(!file||!ready||!session||publishing||busy){excelInput.value='';return;}
    if(leagueDraftPending()){alert('先にリーグ戦の未保存入力を一時保存してください。');excelInput.value='';return;}
    if(!/\.xlsx$/i.test(file.name)){excelStatus.textContent='Excel（.xlsx）ファイルを選択してください。';excelInput.value='';return;}
    const before=Object.fromEntries(Object.keys(labels).map(k=>[k,dataByCategory[k]]));
    const oldAnchors=new Map(leagueAnchors),oldDirty=dirty,oldResume=resumeMatch;
    busy++;excelInput.disabled=true;panel.inert=true;const leagueBox=panel.nextElementSibling;if(leagueBox)leagueBox.inert=true;updatePublishUI();
    try{
        excelStatus.textContent='Excel解析ライブラリを読み込み中…';
        const XLSX=await loadExcelLibrary();
        excelStatus.textContent='Excelの5シートを解析しています（野手・投手を分類）…';
        const workbook=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false,dense:true});
        const parsed={};
        for(const [category,name] of Object.entries(labels)){
            excelStatus.textContent='読み込み中：'+name;
            parsed[category]=readExcelCategory(workbook,category,name);
        }
        const remap=remapLeagueAnchors(before.league,parsed.league);
        if(remap.dropped>0&&!confirm('リーグ戦の既存動画同期点が '+remap.dropped+'件引き継げません。これらの再登録が必要です。新しいExcelで一時保存してもよいですか？（キャンセルすると現在のデータを保持）')){
            throw new Error('同期点が失われないよう、読み込みをキャンセルしました。');
        }
        for(const key of Object.keys(labels))dataByCategory[key]=parsed[key];
        leagueAnchors.clear();invalidateVideoSearchCache();for(const [k,v] of remap.entries)leagueAnchors.set(k,v);
        resumeMatch='';
        await persistDraft();
        markDirty();refreshLeaguePanel();populateFilters();search();
        for(const [cat,name] of Object.entries(labels)){
            const status=uploadStatuses.get(cat);
            status.textContent=parsed[cat].length.toLocaleString('ja-JP')+'件 Excelから一時読み込み済み';status.style.color='#17633a';
        }
        const total=Object.values(parsed).reduce((v,rows)=>v+rows.length,0);
        excelStatus.textContent='5シート一括読み込み完了（未公開）\n'+Object.entries(labels).map(([cat,name])=>name+'：'+parsed[cat].length.toLocaleString('ja-JP')+'件').join('／')+'\n合計：'+total.toLocaleString('ja-JP')+'件。リーグ戦同期点：'+remap.kept+'件引継ぎ'+(remap.dropped?'／要再登録'+remap.dropped+'件':'')+'。最後に「更新終了・一括保存」を押してください。';
        excelStatus.style.color='#17633a';
    }catch(e){
        for(const key of Object.keys(labels))dataByCategory[key]=before[key];
        leagueAnchors.clear();invalidateVideoSearchCache();for(const [k,v] of oldAnchors)leagueAnchors.set(k,v);
        dirty=oldDirty;resumeMatch=oldResume;
        excelStatus.textContent='Excel読み込み失敗：'+e.message+'\n既存のデータ・動画設定は保持されています。';excelStatus.style.color='#b42318';
    }finally{
        excelInput.disabled=false;excelInput.value='';busy--;panel.inert=false;
        if(leagueBox)leagueBox.inert=false;updatePublishUI();
    }
});
if(adminMode)form.before(panel);
// ほかの6項目の条件に合うデータから、各項目の候補を作る。
// 同じ項目で複数選択した値はOR、異なる項目はANDで判定する。
function selectedFilters(){
    return Object.fromEntries(filterFields.map(([key])=>[key,filterSelections[key]]));
}
function matchesFilters(record,selected,exceptKey=null){
    return filterFields.every(([key])=>key===exceptKey||selected[key].size===0||selected[key].has(filterValue(record,key)));
}
function availableFilterValues(all,selected,key){
    const values=new Set();
    for(const record of all){
        if(!matchesFilters(record,selected,key))continue;
        const value=filterValue(record,key);
        if(value)values.add(value);
    }
    if(key==='result')return Object.keys(resultGroups).filter(v=>values.has(v));
    if(key==='category')return Object.keys(labels).filter(v=>values.has(v));
    return [...values].sort((a,b)=>a.localeCompare(b,'ja',{numeric:true}));
}
function filterDisplay(key,value){return key==='category'?(labels[value]||value):value;}
// 現在のデータと他の条件に該当する日付を、月カレンダーに描画する。
// 一度に複数日を選択でき、別の月へ移動しても選択を保持する。
function renderDateCalendar(panel){
    const chosen=filterSelections.date;
    // 動画なしの日付や、現在と異なる野手／投手ビューの日付を候補に出さない。
    const allDates=[...new Set(playableRecordsForView().map(r=>r.date).filter(d=>/^\d{4}\/\d{2}\/\d{2}$/.test(d)))].sort();
    const allDateSet=new Set(allDates);
    const possible=[...panel.available].sort();
    const monthFromDate=d=>d.slice(0,7).replace('/','-');
    if(allDates.length&&!calendarState.hadDates){
        calendarState.month=monthFromDate([...chosen].sort()[0]||possible[possible.length-1]||allDates[allDates.length-1]);
        calendarState.hadDates=true;
    }
    if(!allDates.length)calendarState.hadDates=false;
    if(!calendarState.month){
        calendarState.month=monthFromDate([...chosen].sort()[0]||possible[possible.length-1]||allDates[allDates.length-1]||new Date().toISOString().slice(0,10).replaceAll('-','/'));
    }
    const month=calendarState.month;
    const [y,m]=month.split('-').map(Number);
    if(!y||!m||m<1||m>12)return;
    const firstMonth=allDates.length?monthFromDate(allDates[0]):'';
    const lastMonth=allDates.length?monthFromDate(allDates[allDates.length-1]):'';
    const moveMonth=(delta)=>{
        const dt=new Date(Date.UTC(y,m-1+delta,1));
        const next=dt.getUTCFullYear()+'-'+String(dt.getUTCMonth()+1).padStart(2,'0');
        if((firstMonth&&next<firstMonth)||(lastMonth&&next>lastMonth))return;
        calendarState.month=next;renderDateCalendar(panel);
    };
    const host=document.createDocumentFragment();
    const heading=document.createElement('div');heading.className='bvs-calendar-head';
    const previous=document.createElement('button');previous.type='button';previous.textContent='‹';previous.title='前の月';previous.setAttribute('aria-label','前の月');previous.disabled=!!firstMonth&&month<=firstMonth;
    previous.addEventListener('click',()=>moveMonth(-1));
    const input=document.createElement('input');input.type='month';input.value=month;input.setAttribute('aria-label','表示する年月');
    if(firstMonth)input.min=firstMonth;if(lastMonth)input.max=lastMonth;
    input.addEventListener('change',()=>{if(/^\d{4}-\d{2}$/.test(input.value)){calendarState.month=input.value;renderDateCalendar(panel);}});
    const next=document.createElement('button');next.type='button';next.textContent='›';next.title='次の月';next.setAttribute('aria-label','次の月');next.disabled=!!lastMonth&&month>=lastMonth;
    next.addEventListener('click',()=>moveMonth(1));
    heading.append(previous,input,next);host.append(heading);
    const grid=document.createElement('div');grid.className='bvs-calendar-days';
    for(const day of ['月','火','水','木','金','土','日']){
        const title=document.createElement('div');title.className='bvs-calendar-week';title.textContent=day;grid.append(title);
    }
    const offset=(new Date(Date.UTC(y,m-1,1)).getUTCDay()+6)%7;
    for(let i=0;i<offset;i++)grid.append(document.createElement('span'));
    const days=new Date(Date.UTC(y,m,0)).getUTCDate();
    for(let d=1;d<=days;d++){
        const value=`${y}/${String(m).padStart(2,'0')}/${String(d).padStart(2,'0')}`;
        const day=document.createElement('button');day.type='button';day.textContent=String(d);day.title=value;
        day.setAttribute('aria-label',value);day.setAttribute('aria-pressed',String(chosen.has(value)));
        day.disabled=!allDateSet.has(value);
        if(chosen.has(value))day.classList.add('bvs-calendar-selected');
        else if(!panel.available.has(value))day.classList.add('bvs-calendar-unavailable');
        if(!panel.available.has(value)&&allDateSet.has(value))day.title=value+'（他の絞り込み条件には一致しません）';
        day.addEventListener('click',()=>{if(chosen.has(value))chosen.delete(value);else chosen.add(value);populateFilters();});
        grid.append(day);
    }
    host.append(grid);
    const hint=document.createElement('p');hint.className='bvs-calendar-hint';hint.textContent='色付きの日をクリックして複数日選択できます。薄いグレーは他の条件には一致しない日です。';host.append(hint);
    if(chosen.size){
        const picked=document.createElement('div');picked.className='bvs-calendar-picked';
        for(const date of [...chosen].sort()){
            const chip=document.createElement('button');chip.type='button';chip.textContent=date+' ×';chip.title=date+'の選択を解除';
            chip.addEventListener('click',()=>{chosen.delete(date);populateFilters();});picked.append(chip);
        }
        host.append(picked);
    }
    panel.list.replaceChildren(host);
    panel.meta.textContent=`${panel.available.size}日が他の条件に一致${chosen.size?'・'+chosen.size+'日選択中':''}`;
}
function renderFilterOptions(key){
    const panel=filterPanels[key],selection=filterSelections[key];
    const names=[...selection].map(v=>filterDisplay(key,v));
    panel.summary.textContent=selection.size===0?'すべて':selection.size===1?names[0]:selection.size+'件選択：'+names.slice(0,2).join('、')+(names.length>2?'…':'');
    panel.summary.title=selection.size?names.join('、'):'すべて';
    panel.clear.disabled=selection.size===0;
    if(key==='date'){renderDateCalendar(panel);return;}
    // 他の項目を変更して候補から外れても、チェックした値を勝手に消さない。
    // 選択中の値は下部に残し、利用者が明示的に解除できるようにする。
    const values=[...panel.values,...[...selection].filter(v=>!panel.available.has(v))];
    const query=panel.query.value.trim().toLocaleLowerCase('ja');
    const visible=query?values.filter(v=>filterDisplay(key,v).toLocaleLowerCase('ja').includes(query)):values;
    const scroll=panel.list.scrollTop;
    const fragment=document.createDocumentFragment();
    for(const value of visible){
        const row=document.createElement('label');row.className='bvs-multi-option';
        const check=document.createElement('input');check.type='checkbox';check.checked=selection.has(value);
        const caption=document.createElement('span');
        const noMatch=!panel.available.has(value);
        caption.textContent=filterDisplay(key,value)+(noMatch?'（他の条件とは一致しません）':'');
        if(noMatch)row.classList.add('bvs-unavailable');
        check.addEventListener('change',()=>{
            if(check.checked)filterSelections[key].add(value);
            else filterSelections[key].delete(value);
            populateFilters();
        });
        row.append(check,caption);fragment.append(row);
    }
    if(!visible.length){const p=document.createElement('p');p.className='bvs-multi-meta';p.textContent='候補がありません';fragment.append(p);}
    panel.list.replaceChildren(fragment);
    panel.list.scrollTop=scroll;
    panel.meta.textContent=panel.values.length+'件の候補'+(selection.size?'・'+selection.size+'件選択中':'');
}
function populateFilters(){
    // 絞り込み候補自体も、今のビューで再生可能な投球から生成する。
    const all=playableRecordsForView();
    // Excel/CSVの差し替えで存在しなくなった値だけを解除。組み合わせが0件でも選択は保持する。
    for(const [key] of filterFields){
        const existing=new Set(all.map(r=>filterValue(r,key)).filter(Boolean));
        for(const value of filterSelections[key])if(!existing.has(value))filterSelections[key].delete(value);
    }
    const selected=selectedFilters();
    for(const [key] of filterFields){
        const values=availableFilterValues(all,selected,key),panel=filterPanels[key];
        panel.values=values;panel.available=new Set(values);
        renderFilterOptions(key);
    }
}
function search(){
    const selected=selectedFilters();
    // 動画ありの投球のみ。複数選択・日付カレンダー・2画面の既存動作は変更しない。
    results=playableRecordsForView().filter(r=>matchesFilters(r,selected));
    page=1;clearSelection();renderTable();
}
function renderTable() {
    const size=Number($('pageSize').value)||50,total=Math.ceil(results.length/size);
    page=Math.max(1,Math.min(page,total||1));
    $('resultCount').textContent=results.length.toLocaleString('ja-JP')+'件';
    $('pageInfo').textContent=`${total?page:0} / ${total} ページ`;
    $('prevPage').disabled=page<=1;$('nextPage').disabled=page>=total;
    const body=$('resultBody');body.replaceChildren();
    if(!results.length){const row=body.insertRow(),cell=row.insertCell();cell.colSpan=12;cell.className='empty-table';cell.textContent=activeView==='pitcher'&&!Object.values(dataByCategory).flat().some(r=>recordView(r)==='pitcher')?'投手動画データはまだ公開されていません。管理者サイトで5シートExcelを再読み込みし、「更新終了・一括保存」で公開してください。':!playableRecordsForView().length?'この画面では再生可能な動画がまだ登録されていません。動画URL・再生時刻・リーグ戦の同期点を確認してください。':'条件に一致する再生可能な動画がありません';return;}
    const start=(page-1)*size;
    for(const [i,r] of results.slice(start,start+size).entries()) {
        const tr=body.insertRow();
        const index=start+i;
        tr.classList.toggle('selected',index===selectedIndex);
        tr.classList.toggle('bvs-selected-second',videoCount===2&&index===secondIndex);
        const values=[start+i+1,labels[r.category],r.date,r.opponent,r.batter,r.pitcher,r.inning,r.count,r.pitchType,r.speed===null?'':r.speed+' km/h',r.result];
        for(const value of values){const td=tr.insertCell();td.textContent=value===''?'―':String(value);}
        const cell=tr.insertCell(), actions=document.createElement('div');actions.className='bvs-row-actions';
        const playable=Object.keys(playableAngles(r)).length>0;
        const first=document.createElement('button');first.type='button';first.className='play-button';
        first.textContent=videoCount===2?(playable?'▶ 動画①':'詳細①'):(playable?'▶ 再生':'詳細');
        first.title=playable?'動画①で再生':'動画の再生位置が未登録です';
        first.addEventListener('click',()=>selectVideo(index,true));actions.append(first);
        if(videoCount===2){
            const second=document.createElement('button');second.type='button';second.className='play-button';
            second.textContent=playable?'▶ 動画②':'詳細②';second.title=playable?'動画②で再生':'動画の再生位置が未登録です';
            second.addEventListener('click',()=>selectSecondVideo(index,true));actions.append(second);
        }
        cell.append(actions);
    }
}
$('searchButton').addEventListener('click',search);
form.addEventListener('submit',e=>{e.preventDefault();search();});
// 候補はチェックごとに連動。検索結果は「検索する」で反映する。
function clearAllFilters(){
    for(const [key] of filterFields){
        filterSelections[key].clear();filterPanels[key].query.value='';
        filterPanels[key].details.open=false;
    }
    populateFilters();search();
}
resetButton.addEventListener('click',clearAllFilters);
form.addEventListener('reset',e=>{e.preventDefault();clearAllFilters();});
$('pageSize').addEventListener('change',()=>{page=1;renderTable();});
$('prevPage').addEventListener('click',()=>{page--;renderTable();});
$('nextPage').addEventListener('click',()=>{page++;renderTable();});


// YouTubeのURLは許可したドメイン・動画IDのみ使用する。
function parseTime(value) {
    const v=clean(value);
    if(!v)return null;
    if(/^\d+(?:\.\d+)?$/.test(v)) {const n=Number(v);return Number.isSafeInteger(Math.floor(n))?Math.floor(n):null;}
    if(/^\d+(?::\d{1,2}){1,2}$/.test(v)) {
        const parts=v.split(':').map(Number);
        if(parts.slice(1).some(n=>n>=60))return null;
        const n=parts.reduce((total,n)=>total*60+n,0);return Number.isSafeInteger(n)?n:null;
    }
    const m=v.match(/^(?:(\d+)h)?(?:(\d+)m)?(?:(\d+)s)?$/i);
    if(m&&m.slice(1).some(x=>x!==undefined)) {const n=Number(m[1]||0)*3600+Number(m[2]||0)*60+Number(m[3]||0);return Number.isSafeInteger(n)?n:null;}
    return null;
}
function parseYouTube(angle) {
    if(!angle)return null;
    try {
        const url=new URL(angle.url);
        if(!['https:','http:'].includes(url.protocol))return null;
        const host=url.hostname.toLowerCase(), parts=url.pathname.split('/').filter(Boolean);
        let id='';
        if(host==='youtu.be')id=parts[0];
        else if(['youtube.com','www.youtube.com','m.youtube.com'].includes(host)) {
            if(parts[0]==='watch')id=url.searchParams.get('v');
            else if(['live','embed','shorts'].includes(parts[0]))id=parts[1];
        }
        if(!id||!/^[-_a-zA-Z0-9]{11}$/.test(id))return null;
        const start=parseTime(angle.seconds)??parseTime(url.searchParams.get('t'))??parseTime(url.searchParams.get('start'))??parseTime(new URLSearchParams(url.hash.slice(1)).get('t'));
        // 時刻なしは誤った投球を表示しないよう再生不可とする。
        if(start===null)return null;
        return {id,start};
    }catch{return null;}
}
// STEP 39: 1投球につき、許可されたアングルに有効な動画IDと再生時刻があるか確認。
// 大量データの複数選択でも繰り返しYouTube URLを解析しないようキャッシュする。
// 下書き変更・同期点編集・クラウド読み込み時には必ず破棄する。
let playableAngleCache=new WeakMap();
const playableListCache=new Map();
function invalidateVideoSearchCache(){
    playableAngleCache=new WeakMap();
    playableListCache.clear();
}
function playableAngles(record) {
    const existing=playableAngleCache.get(record);
    if(existing)return existing;
    const allowed=new Set(allowedAngles(recordView(record)));
    const angles=record.category==='league'
        ?Object.fromEntries(angleOrder.filter(a=>allowed.has(a)).map(a=>[a,leagueAngle(record,a)]).filter(([,v])=>v!==null))
        :Object.fromEntries(Object.entries(record.angles||{}).filter(([key])=>allowed.has(key)).map(([key,val])=>[key,parseYouTube(val)]).filter(([,val])=>val!==null));
    playableAngleCache.set(record,angles);
    return angles;
}
function hasPlayableVideo(record){return Object.keys(playableAngles(record)).length>0;}
function playableRecordsForView(view=activeView){
    if(playableListCache.has(view))return playableListCache.get(view);
    const list=Object.values(dataByCategory).flat().filter(r=>recordView(r)===view&&hasPlayableVideo(r));
    playableListCache.set(view,list);
    return list;
}
const angleOrder=['ネット裏','センター裏','一塁側','三塁側','一塁側(投手)','三塁側(投手)'];
const angleButtons=[...document.querySelectorAll('.angle-button')];
angleButtons.forEach((button,i)=>{if(angleOrder[i]==='センター裏')button.textContent=angleLabel(angleOrder[i]);});
// STEP 34: 動画①は従来の操作系を維持。動画②を独立したプレイヤーとして追加する。
const player=$('videoPlayer');
let videoCount=1,secondIndex=-1,secondAngle=null;
const videoLink=document.createElement('a');
videoLink.textContent='YouTubeで開く ↗';videoLink.target='_blank';videoLink.rel='noopener';videoLink.hidden=true;
videoLink.style.cssText='display:none;color:#24558e;font-size:13px;margin-top:10px;text-align:right;';
player.after(videoLink);
const videoStage=document.createElement('div');videoStage.className='bvs-video-stage';videoStage.dataset.count='1';
const modeBar=document.createElement('div');modeBar.className='bvs-mode-bar';
const modeTitle=document.createElement('strong');modeTitle.textContent='動画表示：';
const oneMode=document.createElement('button');oneMode.type='button';oneMode.className='bvs-mode-button active';oneMode.textContent='1画面';oneMode.setAttribute('aria-pressed','true');
const twoMode=document.createElement('button');twoMode.type='button';twoMode.className='bvs-mode-button';twoMode.textContent='2画面';twoMode.setAttribute('aria-pressed','false');
const modeExplain=document.createElement('span');modeExplain.className='bvs-mode-explain';modeExplain.textContent='';
modeBar.append(modeTitle,oneMode,twoMode,modeExplain);
player.before(modeBar,videoStage);
const firstSlot=document.createElement('div');firstSlot.className='bvs-dual-slot';
const firstTitle=document.createElement('div');firstTitle.className='bvs-dual-title';firstTitle.textContent='動画①';
// 従来のアングルボタンを動画①の直下に移動。クリック処理・選択状態は引き続き従来のものを使う。
const originalAngleHost=angleButtons[0]?.parentElement||null;
const firstAngleTitle=document.createElement('div');firstAngleTitle.className='bvs-dual-angle-title';firstAngleTitle.textContent='動画①の撮影アングル';
const firstAngleList=document.createElement('div');firstAngleList.className='bvs-dual-angle-list bvs-first-angle-list';
firstAngleList.setAttribute('role','group');firstAngleList.setAttribute('aria-label','動画①の撮影アングル');
for(const button of angleButtons){button.type='button';firstAngleList.append(button);}
// 元のボタンの空になった入れ物だけ隠し、他の再生操作は維持する。
if(originalAngleHost&&!originalAngleHost.querySelector('button,select,input,a'))originalAngleHost.style.display='none';
firstSlot.append(firstTitle,player,firstAngleTitle,firstAngleList,videoLink);videoStage.append(firstSlot);
const secondSlot=document.createElement('div');secondSlot.className='bvs-dual-slot';secondSlot.style.display='none';
const secondTitle=document.createElement('div');secondTitle.className='bvs-dual-title';secondTitle.textContent='動画②';
const secondPlayer=document.createElement('div');secondPlayer.id='bvsSecondPlayer';secondPlayer.className=player.className+' bvs-second-player';
const secondInfo=document.createElement('p');secondInfo.className='bvs-second-detail';secondInfo.textContent='比較したい投球を検索結果から選んでください';
const secondActions=document.createElement('div');secondActions.className='bvs-second-actions';
const secondPrev=document.createElement('button');secondPrev.type='button';secondPrev.textContent='◀ 前の動画②';
const secondPosition=document.createElement('span');secondPosition.style.cssText='font-size:12px;color:#64748b';secondPosition.textContent='0 / 0';
const secondNext=document.createElement('button');secondNext.type='button';secondNext.textContent='次の動画② ▶';
const copyToSecond=document.createElement('button');copyToSecond.type='button';copyToSecond.textContent='動画①の投球を表示';copyToSecond.title='同じ投球を別アングルで比較できます';
secondActions.append(secondPrev,secondPosition,secondNext,copyToSecond);
const secondAngleTitle=document.createElement('div');secondAngleTitle.className='bvs-dual-angle-title';secondAngleTitle.textContent='動画②の撮影アングル';
const secondAngleList=document.createElement('div');secondAngleList.className='bvs-dual-angle-list bvs-second-angle-list';
secondAngleList.setAttribute('role','group');secondAngleList.setAttribute('aria-label','動画②の撮影アングル');
const secondAngleButtons=angleOrder.map(a=>{
    const button=document.createElement('button');button.type='button';button.className='bvs-compare-angle';button.textContent=angleLabel(a);
    button.addEventListener('click',()=>displaySecondAngle(a));secondAngleList.append(button);return button;
});
const secondLink=document.createElement('a');secondLink.className='bvs-second-link';secondLink.textContent='動画②をYouTubeで開く ↗';
secondLink.target='_blank';secondLink.rel='noopener';secondLink.hidden=true;
secondSlot.append(secondTitle,secondPlayer,secondInfo,secondActions,secondAngleTitle,secondAngleList,secondLink);
videoStage.append(secondSlot);
const showInfo=(r)=>{
    const values={infoBatter:r?.batter,infoPitcher:r?.pitcher,infoPitchType:r?.pitchType,infoSpeed:r?.speed!=null?r.speed+' km/h':'',infoCount:r?.count,infoResult:r?.result};
    for(const [id,value] of Object.entries(values))$(id).textContent=value||'―';
};
function makePlaceholder(host,message){
    const div=document.createElement('div');div.className='video-placeholder';div.style.padding='16px';
    const icon=document.createElement('div');icon.className='play-icon';icon.textContent='▶';
    const p=document.createElement('p');p.textContent=message;div.append(icon,p);host.replaceChildren(div);
}
function placeholder(message){
    makePlaceholder(player,message);
    videoLink.hidden=true;videoLink.style.display='none';videoLink.removeAttribute('href');
}
function secondPlaceholder(message){
    makePlaceholder(secondPlayer,message);
    secondLink.hidden=true;secondLink.removeAttribute('href');
}
function buildPlayer(host,link,r,key){
    const angles=playableAngles(r),video=angles[key];if(!video)return false;
    const watch=new URL('https://www.youtube.com/watch');watch.searchParams.set('v',video.id);watch.searchParams.set('t',video.start+'s');
    link.href=watch.href;link.hidden=false;if(link===videoLink)link.style.display='block';
    if(location.protocol==='file:'){
        makePlaceholder(host,'サイト内再生にはローカルサーバーで開いてください。YouTubeのリンクからは再生できます。');
        return true;
    }
    const src=new URL('https://www.youtube.com/embed/'+video.id);
    src.searchParams.set('start',String(video.start));src.searchParams.set('autoplay','1');src.searchParams.set('playsinline','1');
    const iframe=document.createElement('iframe');iframe.title=(activeView==='pitcher'?r.pitcher:r.batter)+'｜'+angleLabel(key)+'｜'+(activeView==='pitcher'?'投手動画':'野手動画');
    iframe.allow='accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share';
    iframe.allowFullscreen=true;iframe.referrerPolicy='strict-origin-when-cross-origin';iframe.src=src.href;
    host.replaceChildren(iframe);return true;
}
function updateSecondControls(){
    const record=results[secondIndex],angles=record?playableAngles(record):{};
    secondPosition.textContent=record?`${secondIndex+1} / ${results.length}`:`0 / ${results.length}`;
    secondPrev.disabled=!record||secondIndex===0;secondNext.disabled=!record||secondIndex===results.length-1;
    copyToSecond.disabled=selectedIndex<0;
    secondAngleButtons.forEach((b,i)=>{
        const a=angleOrder[i];b.style.display=allowedAngles(activeView).includes(a)?'':'none';b.disabled=!angles[a];b.classList.toggle('active',a===secondAngle&&!!angles[a]);b.setAttribute('aria-pressed',String(a===secondAngle&&!!angles[a]));
    });
    secondInfo.textContent=record?`${record.date}｜${labels[record.category]}｜打者：${record.batter}｜投手：${record.pitcher}｜${record.inning||'―'}｜${record.result||'―'}`:'比較したい投球を検索結果から選んでください';
}
function clearSelection(){
    selectedIndex=-1;selectedAngle=null;showInfo(null);placeholder('動画を選択してください');
    $('prevVideo').disabled=true;$('nextVideo').disabled=true;$('videoPosition').textContent='0 / '+results.length;
    angleButtons.forEach((b,i)=>{b.style.display=allowedAngles(activeView).includes(angleOrder[i])?'':'none';b.disabled=true;b.classList.remove('active');b.setAttribute('aria-pressed','false');});
    secondIndex=-1;secondAngle=null;secondPlaceholder('比較したい動画を選択してください');updateSecondControls();
}
function displayAngle(key){
    const r=results[selectedIndex];if(!r)return;
    const angles=playableAngles(r),video=angles[key];if(!video)return;
    selectedAngle=key;
    angleButtons.forEach((b,i)=>{b.style.display=allowedAngles(activeView).includes(angleOrder[i])?'':'none';b.disabled=!angles[angleOrder[i]];b.classList.toggle('active',angleOrder[i]===key);b.setAttribute('aria-pressed',String(angleOrder[i]===key));});
    buildPlayer(player,videoLink,r,key);
}
function displaySecondAngle(key){
    const r=results[secondIndex];if(!r)return;
    if(!playableAngles(r)[key])return;
    secondAngle=key;buildPlayer(secondPlayer,secondLink,r,key);updateSecondControls();
}
function selectVideo(index,scroll=false){
    if(index<0||index>=results.length)return;
    selectedIndex=index;const r=results[index],angles=playableAngles(r);
    showInfo(r);$('videoPosition').textContent=`${index+1} / ${results.length}`;
    $('prevVideo').disabled=index===0;$('nextVideo').disabled=index===results.length-1;
    const preferred=angles[selectedAngle]?selectedAngle:angleOrder.find(a=>angles[a]);
    if(preferred){displayAngle(preferred);}else{
        placeholder(r.category==='league'?'リーグ戦：動画の基準時刻が未登録、または撮影中断区間です。':'この投球の動画URLまたは再生時刻が未登録です。');
        angleButtons.forEach((b,i)=>{b.style.display=allowedAngles(activeView).includes(angleOrder[i])?'':'none';b.disabled=true;b.classList.remove('active');b.setAttribute('aria-pressed','false');});
    }
    page=Math.floor(index/(Number($('pageSize').value)||50))+1;renderTable();
    updateSecondControls();
    if(scroll)videoStage.scrollIntoView({behavior:'smooth',block:'start'});
}
function selectSecondVideo(index,scroll=false,preferDifferent=false){
    if(index<0||index>=results.length)return;
    secondIndex=index;const r=results[index],angles=playableAngles(r);
    let preferred=angles[secondAngle]?secondAngle:angleOrder.find(a=>angles[a]);
    if(preferDifferent&&index===selectedIndex){const alternate=angleOrder.find(a=>a!==selectedAngle&&angles[a]);if(alternate)preferred=alternate;}
    if(preferred)displaySecondAngle(preferred);
    else secondPlaceholder(r.category==='league'?'リーグ戦：動画の基準時刻が未登録、または撮影中断区間です。':'この投球の動画URLまたは再生時刻が未登録です。');
    updateSecondControls();
    page=Math.floor(index/(Number($('pageSize').value)||50))+1;renderTable();
    if(scroll)videoStage.scrollIntoView({behavior:'smooth',block:'start'});
}
function setVideoCount(count){
    if(count!==1&&count!==2)return;
    videoCount=count;videoStage.dataset.count=String(count);secondSlot.style.display=count===2?'':'none';
    oneMode.classList.toggle('active',count===1);twoMode.classList.toggle('active',count===2);
    oneMode.setAttribute('aria-pressed',String(count===1));twoMode.setAttribute('aria-pressed',String(count===2));
    modeExplain.textContent=count===2?'左右の動画それぞれの直下でアングルを選択できます。検索結果から別々の投球も選べます。':'';
    if(count===1){
        // 画面から隠す際には右プレイヤーを破棄して音声の再生も止める。
        secondPlaceholder('比較したい動画を選択してください');
    }else{
        if(secondIndex>=0&&results[secondIndex]){
            const r=results[secondIndex],a=playableAngles(r);
            const preferred=a[secondAngle]?secondAngle:angleOrder.find(k=>a[k]);
            if(preferred)displaySecondAngle(preferred);
        }else if(selectedIndex>=0){
            // 同じ投球に別角度があれば、2画面切替時にその角度を表示する。
            const angles=playableAngles(results[selectedIndex]);
            const different=angleOrder.find(a=>a!==selectedAngle&&angles[a]);
            if(different){secondAngle=different;selectSecondVideo(selectedIndex,false,true);}
        }
    }
    updateSecondControls();renderTable();
}
oneMode.addEventListener('click',()=>setVideoCount(1));
twoMode.addEventListener('click',()=>setVideoCount(2));
angleButtons.forEach((b,i)=>b.addEventListener('click',()=>displayAngle(angleOrder[i])));
$('prevVideo').addEventListener('click',()=>selectVideo(selectedIndex-1,true));
$('nextVideo').addEventListener('click',()=>selectVideo(selectedIndex+1,true));
secondPrev.addEventListener('click',()=>selectSecondVideo(secondIndex-1,true));
secondNext.addEventListener('click',()=>selectSecondVideo(secondIndex+1,true));
copyToSecond.addEventListener('click',()=>{if(selectedIndex>=0)selectSecondVideo(selectedIndex,true,true);});
clearSelection();

// STEP 8/30: リーグ戦の同期。登録はブラウザの一時保存に保持。
const leagueAnchors = new Map();
const leagueMatchKey = r => JSON.stringify([r.date,r.opponent]);
const leagueGroupKey = r => recordView(r)==='pitcher'?JSON.stringify(['pitcher',r.date,r.opponent,r.inning]):JSON.stringify([r.date,r.opponent,r.inning]);
function clockSeconds(value) {
    const m=clean(value).match(/^\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}\s+(\d{1,2}:\d{2}:\d{2})$/);
    return m?parseTime(m[1]):null;
}
function chapterLines(text) {
    const found=new Map(), ignored=[];
    for(const raw of text.split(/\r?\n/)) {
        const line=raw.normalize('NFKC').trim();if(!line)continue;
        const m=line.match(/^(\d+(?::\d{1,2}){1,2})\s*[-–—：:　 ]*\s*(?:第\s*)?(\d+)\s*(?:回\s*)?(表|裏)(?:\s|$|[^\d])/);
        if(!m||parseTime(m[1])===null){ignored.push(line);continue;}
        const inning=Number(m[2])+'回'+m[3];
        if(found.has(inning))throw new Error(inning+' のチャプターが重複しています。');
        found.set(inning,parseTime(m[1]));
    }
    return {found,ignored};
}
function anchorSlot(r, angle) {return JSON.stringify([leagueGroupKey(r),angle]);}
function putAnchor(r, angle, anchor) {
    const key=anchorSlot(r,angle), list=leagueAnchors.get(key)||[];
    leagueAnchors.set(key,[...list.filter(a=>a.row!==r.sourceRow),{...anchor,row:r.sourceRow,clock:clockSeconds(r.timestamp)}].sort((a,b)=>a.row-b.row));
    invalidateVideoSearchCache();
}
function leagueAngle(r,angle) {
    const list=leagueAnchors.get(anchorSlot(r,angle))||[];
    let a=null;for(const item of list){if(item.row>r.sourceRow)break;a=item;}
    const clock=clockSeconds(r.timestamp);
    if(!a||a.missing||clock===null||a.clock===null||clock<a.clock)return null;
    const nominal=a.start+clock-a.clock;
    // 3秒の早出し。撮影再開地点は別動画の先頭より前にしない。
    return {id:a.id,start:Math.max(0,nominal-3),nominal};
}
// STEP 30: 2アングルを試合単位で登録し、ローカル下書きに保存。
const leagueRegistrationAngles=['ネット裏','センター裏'];
function planLeagueAngle(gameRows, angle, input) {
    const url=clean(input.url), raw=clean(input.raw), corrected=clean(input.corrected);
    if(!url&&!raw&&!corrected)return [];
    const name=angleLabel(angle);
    if(!url||!raw||!corrected)throw new Error(name+'：3項目を入力してください（登録しない場合は3項目とも空欄）。');
    const video=parseYouTube({url,seconds:'0'});
    if(!video)throw new Error(name+'：正しいYouTube URLを入力してください。');
    const original=chapterLines(raw);
    const first=new Map();for(const r of gameRows)if(r.inning&&!first.has(r.inning))first.set(r.inning,r);
    const innings=[...original.found.keys()].filter(i=>first.has(i));
    if(!innings.length)throw new Error(name+'：元チャプターに慶應攻撃回がありません。例：05:00 1回表');
    const lines=corrected.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
    const timesOnly=lines.every(x=>/^\d+(?::\d{1,2}){1,2}$/.test(x)&&parseTime(x)!==null);
    if(timesOnly&&lines.length!==innings.length)throw new Error(name+'：時刻だけの場合は慶應の攻撃回 '+innings.length+' 行分を元チャプター順で入力してください。');
    const fixed=timesOnly?{found:new Map(innings.map((i,j)=>[i,parseTime(lines[j])])),ignored:[]}:chapterLines(corrected);
    if(fixed.ignored.length||!fixed.found.size)throw new Error(name+'：修正時刻は「05:02 1回表」の形式で1行ずつ入力してください。');
    const absent=innings.filter(i=>!fixed.found.has(i));
    const extra=[...fixed.found.keys()].filter(i=>!original.found.has(i));
    if(absent.length)throw new Error(name+'：修正時刻が不足しています：'+absent.join('、'));
    if(extra.length)throw new Error(name+'：元チャプターにない修正行：'+extra.join('、'));
    const planned=[];
    for(const inning of innings){
        const r=first.get(inning), start=fixed.found.get(inning);
        if(clockSeconds(r.timestamp)===null)throw new Error(name+'：'+inning+' のCSV時刻が不正です。');
        planned.push({r,angle,anchor:{id:video.id,start,originalStart:original.found.get(inning),originalChapterText:raw,registrationUrl:url}});
    }
    return planned;
}
function setupLeaguePanel(view) {
    const registrationAngles=allowedAngles(view).filter(a=>view==='pitcher'||a==='ネット裏'||a==='センター裏');
    const box=document.createElement('details');box.style.cssText=panel.style.cssText;
    const summary=document.createElement('summary');summary.textContent='リーグ戦 '+(view==='pitcher'?'投手':'野手')+'動画登録（1試合ずつ）';summary.style.cssText='font-weight:bold;cursor:pointer;margin-bottom:8px';box.append(summary);
    const p=(parent,t)=>{const el=document.createElement('p');el.textContent=t;el.style.cssText='margin:8px 0;line-height:1.5;overflow-wrap:anywhere;';parent.append(el);return el;};
    p(box,'試合を選択し、各アングルの元チャプター・投球動作開始2秒前の修正時刻・YouTube URLを登録してください。再生は修正時刻のさらに3秒前から。撮影がないアングルの3項目は空欄のままにしてください。各試合を一時保存した後、最後に更新管理欄の「更新終了」を押すとまとめて公開されます。');
    function field(parent,label,el){const l=document.createElement('label');l.textContent=label;l.style.cssText='display:block;font-weight:bold;margin:10px 0 4px';el.style.cssText='display:block;width:100%;margin-top:4px;padding:6px;border:1px solid #b9c5d3;border-radius:4px;font-size:12px;box-sizing:border-box;';l.append(el);parent.append(l);return el;}
    const match=field(box,'試合',document.createElement('select'));
    const fields=new Map();let edited=false, currentMatch='';leagueDraftChecks.push(()=>edited);
    const progress=p(box,'');
    for(const a of registrationAngles){
        const section=document.createElement('section');section.style.cssText='border-top:1px solid #cdd7e3;margin-top:14px;padding-top:10px;';box.append(section);
        const title=document.createElement('h3');title.textContent=angleLabel(a);section.append(title);
        const raw=field(section,'① チャプター（YouTubeからコピー）',document.createElement('textarea'));raw.rows=5;raw.placeholder='00:00 試合開始\n05:00 1回表\n18:25 1回裏';
        const corrected=field(section,'② 投球動作開始の2秒前（動画内時刻）',document.createElement('textarea'));corrected.rows=5;corrected.placeholder='05:02 1回表\n18:27 1回裏';
        const url=field(section,'③ YouTube URL',document.createElement('input'));url.type='url';url.placeholder='https://youtu.be/...';
        const info=p(section,'');fields.set(a,{url,raw,corrected,info});
        for(const el of [url,raw,corrected])el.addEventListener('input',()=>{edited=true;info.textContent='未保存の入力あり';updatePublishUI();});
    }
    const status=p(box,'試合を選択してください。');status.setAttribute('aria-live','polite');
    const save=document.createElement('button');save.type='button';save.className='search-button';save.textContent='一時保存して次の試合へ';box.append(save);
    const prev=document.createElement('button');prev.type='button';prev.className='reset-button';prev.textContent='前の試合へ';box.append(prev);
    const skip=document.createElement('button');skip.type='button';skip.className='reset-button';skip.textContent='次の試合へ（変更なし）';box.append(skip);
    function neighbor(delta){const opts=[...match.options].filter(o=>o.value);const i=opts.findIndex(o=>o.value===match.value);return opts[i+delta]?.value||'';}
    function updateProgress(){const keys=new Set(dataByCategory.league.filter(r=>recordView(r)===view).map(leagueMatchKey)), registered=new Set(dataByCategory.league.filter(r=>recordView(r)===view&&registrationAngles.some(a=>leagueAngle(r,a))).map(leagueMatchKey));progress.textContent='再生位置登録済み（片方以上）：'+registered.size+' / '+keys.size+'試合｜未公開の一時保存は最後に更新終了';}
    const rows=()=>dataByCategory.league.filter(r=>recordView(r)===view&&leagueMatchKey(r)===match.value);
    const requireMatch=()=>{if(!rows().length)throw new Error('リーグ戦CSVと試合を選択してください。');};
    const count=()=>rows().filter(r=>registrationAngles.some(a=>leagueAngle(r,a))).length;
    const say=(msg,error=false)=>{status.textContent=msg;status.style.color=error?'#b42318':'#17633a';};
    const firstRows=()=>{const m=new Map();for(const r of rows())if(r.inning&&!m.has(r.inning))m.set(r.inning,r);return [...m.values()];};
    function populateInputs(){
        currentMatch=match.value;edited=false;url.value='';time.value='';inning.value='';pitch.value='';
        for(const [a,f] of fields){
            const entries=firstRows().map(r=>({r,anchor:(leagueAnchors.get(anchorSlot(r,a))||[]).find(x=>x.row===r.sourceRow&&!x.missing)})).filter(x=>x.anchor);
            // 複数動画の既存登録は単一URL欄に誤って集約しない。
            const ids=new Set(entries.map(x=>x.anchor.id));
            f.url.value=ids.size===1?(entries[0].anchor.registrationUrl||'https://www.youtube.com/watch?v='+entries[0].anchor.id):'';
            const fmt=(s)=>[Math.floor(s/3600),Math.floor(s/60)%60,s%60].map((v,i)=>i===0?String(v):String(v).padStart(2,'0')).join(':');
            const canFill=ids.size===1;
            f.raw.value=canFill?(entries.find(x=>x.anchor.originalChapterText)?.anchor.originalChapterText||entries.map(({r,anchor})=>fmt(anchor.originalStart??anchor.start)+' '+r.inning).join('\n')):'';
            // 元テキストと異なるイニングが過去に追加された場合も補完。
            if(canFill){const known=chapterLines(f.raw.value).found;for(const {r,anchor} of entries)if(!known.has(r.inning))f.raw.value+='\n'+fmt(anchor.originalStart??anchor.start)+' '+r.inning;}
            f.corrected.value=canFill?entries.map(({r,anchor})=>fmt(anchor.start)+' '+r.inning).join('\n'):'';
            f.initial=JSON.stringify([f.url.value,f.raw.value,f.corrected.value]);
            f.info.textContent=ids.size>1?'複数動画の既存登録あり。空欄のままなら維持します。途中補正から編集可能。':entries.length?entries.length+'イニング登録済み'+(entries.some(x=>x.anchor.originalStart===undefined)?'（旧登録は元チャプターの代わりに保存時刻を表示）':''):'未登録';
        }
        save.disabled=!match.value;prev.disabled=!neighbor(-1);skip.disabled=!neighbor(1);
        updateInnings();updateProgress();
    }
    function updateMatches(preferred=match.value){
        const previous=preferred;const m=new Map(dataByCategory.league.filter(r=>recordView(r)===view).map(r=>[leagueMatchKey(r),r.date+' vs '+r.opponent]));
        match.replaceChildren(new Option('選択してください',''),...[...m].sort((x,y)=>x[1].localeCompare(y[1])).map(([k,label])=>new Option(label,k)));
        if(m.has(previous))match.value=previous;
        populateInputs();say(match.value?'この試合の再生位置登録済み：'+count()+' / '+rows().length+'件':'試合を選択してください。');
    }
    match.addEventListener('change',()=>{
        if(busy||publishing){match.value=currentMatch;return;}
        if(edited&&!confirm('未保存の入力を破棄して別の試合を選択しますか？')){match.value=currentMatch;return;}
        populateInputs();updatePublishUI();say('この試合の再生位置登録済み：'+count()+' / '+rows().length+'件');
    });
    function move(delta){if(!ready||busy||publishing)return;const target=neighbor(delta);if(!target)return;if(edited&&!confirm('未保存の入力を破棄して移動しますか？'))return;match.value=target;populateInputs();updatePublishUI();say('この試合の再生位置登録済み：'+count()+' / '+rows().length+'件');}
    prev.addEventListener('click',()=>move(-1));skip.addEventListener('click',()=>move(1));
    save.addEventListener('click',async()=>{
        if(!ready||!session||publishing||busy)return;
        try{
            requireMatch();
            const plan=[...fields].flatMap(([a,f])=>JSON.stringify([f.url.value,f.raw.value,f.corrected.value])===f.initial?[]:planLeagueAngle(rows(),a,{url:f.url.value,raw:f.raw.value,corrected:f.corrected.value}));
            if(!plan.length)throw new Error('保存する変更がありません。少なくとも1アングルの3項目を入力・変更してください。');
            const label=match.selectedOptions[0].textContent;
            const before=new Map(leagueAnchors),oldDirty=dirty,oldResume=resumeMatch;
            const next=neighbor(1);
            busy++;box.inert=true;updatePublishUI();
            try{
                for(const {r,angle,anchor} of plan)putAnchor(r,angle,anchor);
                resumeMatch=next||match.value;
                await persistDraft();
            }catch(e){leagueAnchors.clear();invalidateVideoSearchCache();for(const [k,v] of before)leagueAnchors.set(k,v);dirty=oldDirty;resumeMatch=oldResume;say('一時保存失敗：'+e.message+'。入力は保持しています。',true);return;}
            finally{busy--;box.inert=false;updatePublishUI();}
            markDirty();const savedCount=count(),total=rows().length;
            if(next)match.value=next;
            populateInputs();search();updatePublishUI();
            say(`${label}：一時保存完了（再生位置 ${savedCount} / ${total}件）。`+(next?' 次の試合を入力できます。':' 最後の試合です。更新終了で公開してください。'));
        }catch(e){say(e.message,true);}
    });
    // 例外対応は折りたたみ内に残す。通常登録は上の入力欄を使います。
    const advanced=document.createElement('details');advanced.style.cssText='margin-top:16px;border-top:1px solid #cdd7e3;padding-top:10px;';box.append(advanced);
    const as=document.createElement('summary');as.textContent='途中補正・撮影中断（必要な場合のみ）';advanced.append(as);
    p(advanced,'上の入力項目に未保存入力がある場合は先に試合を保存してください。中断した最初の投球に「動画なし」、再開した最初の投球に新しい基準時刻を登録。補正も一時保存され、下の更新終了でまとめて公開します。');
    const angle=field(advanced,'撮影アングル',document.createElement('select'));angle.replaceChildren(...registrationAngles.map(a=>new Option(angleLabel(a),a)));
    const url=field(advanced,'再開動画のYouTube URL（空欄なら上のURL）',document.createElement('input'));url.type='url';
    const inning=field(advanced,'イニング',document.createElement('select'));
    const pitch=field(advanced,'基準にする投球（CSV順）',document.createElement('select'));
    const time=field(advanced,'動画内の時刻（投球動作の2秒前）',document.createElement('input'));time.placeholder='1:23:45';
    function updatePitches(){const previous=pitch.value;const a=rows().filter(r=>r.inning===inning.value);pitch.replaceChildren(new Option('選択してください',''),...a.map((r,i)=>{const item=(leagueAnchors.get(anchorSlot(r,angle.value))||[]).find(x=>x.row===r.sourceRow);return new Option((i+1)+'球目｜'+r.timestamp.slice(-8)+'｜'+r.batter+'｜打席内'+r.pitchInAtBat+'球'+(item?item.missing?'【動画なし】':'【基準登録済】':''),String(r.sourceRow));}));if([...pitch.options].some(o=>o.value===previous))pitch.value=previous;}
    function updateInnings(){const previous=inning.value;inning.replaceChildren(new Option('選択してください',''),...[...new Set(rows().map(r=>r.inning).filter(Boolean))].map(i=>new Option(i,i)));if([...inning.options].some(o=>o.value===previous))inning.value=previous;updatePitches();}
    function button(label,action){const b=document.createElement('button');b.type='button';b.textContent=label;b.className='search-button';b.style.cssText='margin:6px 0;font-size:12px;padding:8px;';b.addEventListener('click',async()=>{if(!ready||publishing||busy)return;if(edited){say('先に上の入力項目を一時保存してください。',true);return;}const before=new Map(leagueAnchors),oldDirty=dirty,oldResume=resumeMatch;busy++;box.inert=true;updatePublishUI();try{action();resumeMatch=match.value;await persistDraft();markDirty();search();populateInputs();say('途中補正を一時保存しました。下の「更新終了」で公開してください。');}catch(e){leagueAnchors.clear();invalidateVideoSearchCache();for(const [k,v] of before)leagueAnchors.set(k,v);dirty=oldDirty;resumeMatch=oldResume;say(e.message,true);}finally{busy--;box.inert=false;updatePublishUI();}});advanced.append(b);}
    const chosen=()=>{requireMatch();const r=rows().find(r=>r.inning===inning.value&&String(r.sourceRow)===pitch.value);if(!r)throw new Error('基準にする投球を選択してください。');return r;};
    inning.addEventListener('change',updatePitches);angle.addEventListener('change',updatePitches);
    button('この投球から再同期',()=>{const r=chosen(),v=parseYouTube({url:url.value||fields.get(angle.value).url.value,seconds:'0'}),start=parseTime(time.value);if(!v||start===null||clockSeconds(r.timestamp)===null)throw new Error('動画URL・動画内時刻・CSV時刻を確認してください。');putAnchor(r,angle.value,{id:v.id,start});});
    button('この投球から動画なし',()=>putAnchor(chosen(),angle.value,{missing:true}));
    button('選択した投球の基準・中断点を削除',()=>{const r=chosen(),key=anchorSlot(r,angle.value);leagueAnchors.set(key,(leagueAnchors.get(key)||[]).filter(a=>a.row!==r.sourceRow));});
    panel.after(box);box.dataset.view=view;leaguePanels.push(box);box.hidden=true;updateMatches();return updateMatches;
}
const refreshLeaguePanel=adminMode?(()=>{const pitcher=setupLeaguePanel('pitcher'),batter=setupLeaguePanel('batter');return (preferred)=>{pitcher(preferred);batter(preferred);};})():()=>{};
leagueDraftPending=()=>leagueDraftChecks.some(fn=>fn());
populateFilters();


// STEP 16: CSV差し替え時に行番号だけで同期点を引き継がない。
function remapLeagueAnchors(oldRows,newRows){
    const signature=r=>JSON.stringify([recordView(r),r.date,r.opponent,r.inning,r.timestamp,r.batter,r.pitcher,r.pitchInAtBat,r.count]);
    const index=rows=>{const m=new Map();for(const r of rows){const k=signature(r);if(m.has(k))m.set(k,null);else m.set(k,r);}return m;};
    const oldByRow=new Map(oldRows.map(r=>[recordView(r)+':'+r.sourceRow,r]));
    const oldIndex=index(oldRows),newIndex=index(newRows),entries=new Map();
    let kept=0,dropped=0;
    for(const [key,list] of leagueAnchors){
        const [group,angle]=JSON.parse(key);
        const role=JSON.parse(group)[0]==='pitcher'?'pitcher':'batter';
        const pending=[];let failed=false;
        for(const a of list){
            const old=oldByRow.get(role+':'+a.row),sig=old&&signature(old),next=sig&&newIndex.get(sig);
            if(!old||!oldIndex.get(sig)||!next||leagueGroupKey(old)!==group||clockSeconds(next.timestamp)!==a.clock){failed=true;break;}
            pending.push([anchorSlot(next,angle),{...a,row:next.sourceRow}]);
        }
        // 中断点だけが消えて誤った映像に繋がらないよう、曖昧な組は全て再登録。
        if(failed){dropped+=list.length;continue;}
        for(const [slot,a] of pending){if(!entries.has(slot))entries.set(slot,[]);entries.get(slot).push(a);kept++;}
    }
    for(const list of entries.values())list.sort((a,b)=>a.row-b.row);
    return {entries,kept,dropped};
}

// IndexedDBに確定版を一つのトランザクションで保存する。
const dbName='baseball-video-search-step16';
function openStore(){return new Promise((resolve,reject)=>{
    const request=indexedDB.open(dbName,1);
    request.onupgradeneeded=()=>request.result.createObjectStore('snapshots');
    request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);
    request.onblocked=()=>reject(new Error('他のタブを閉じて再試行してください。'));
});}
async function readPublished(key='published'){const db=await openStore();try{return await new Promise((resolve,reject)=>{
    const tx=db.transaction('snapshots','readonly'),req=tx.objectStore('snapshots').get(key);
    req.onsuccess=()=>resolve(req.result||null);req.onerror=()=>reject(req.error);tx.onabort=()=>reject(tx.error);
});}finally{db.close();}}
async function savePublished(snapshot,key='published'){const db=await openStore();try{await new Promise((resolve,reject)=>{
    const tx=db.transaction('snapshots','readwrite');tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error||new Error('保存が中断されました。'));
    tx.objectStore('snapshots').put(snapshot,key);
});}finally{db.close();}}
// STEP 22: Supabase 接続。公開用キーだけをここに貼り付ける。
const SUPABASE_URL = 'https://ftiuxochwailzqowhqvn.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_ompcWMNKOLV69g4sXeN8uw_RcXdacBA';
const ADMIN_UID = '3c040525-96d4-4e38-99cd-db83a17e2fa2';
let session=null, sessionExpires=0, cloudVersion=null;
// 下書きは認証情報を含まない。公開版と別のキーへ保存し、同じ公開版からのみ復元する。
const draftKey='league-batch-draft-v1:'+SUPABASE_URL+':'+ADMIN_UID;
async function persistDraft(){if(!session||!ready)throw new Error('管理者ログインが必要です。');await savePublished({schema:1,baseVersion:cloudVersion,nextMatch:resumeMatch,snapshot:{version:1,publishedAt:lastPublished,data:dataByCategory,anchors:[...leagueAnchors]}},draftKey);}
const configured=()=>/^sb_publishable_[A-Za-z0-9_-]+$/.test(SUPABASE_PUBLISHABLE_KEY);
async function request(path,{method='GET',body,token,prefer}={}) {
    if(!configured())throw new Error('script.js の Publishable key を設定してください。');
    const headers={apikey:SUPABASE_PUBLISHABLE_KEY};
    if(token)headers.Authorization='Bearer '+token;
    if(body!==undefined)headers['Content-Type']='application/json';
    if(prefer)headers.Prefer=prefer;
    const response=await fetch(SUPABASE_URL+path,{method,headers,body:body===undefined?undefined:JSON.stringify(body),cache:'no-store'});
    const raw=await response.text();let value=null;try{value=raw?JSON.parse(raw):null;}catch{value=raw;}
    if(!response.ok)throw new Error(`Supabase ${response.status}：${value?.message||value?.msg||value?.error_description||value?.error||'通信エラー'}`);
    return value;
}
function setSession(value){
    if(value?.user?.id!==ADMIN_UID||!value.access_token||!value.refresh_token)throw new Error('管理者アカウントではありません。');
    session=value;sessionExpires=Date.now()+Number(value.expires_in||3600)*1000;
}
async function adminToken(){
    if(!session)throw new Error('管理者ログインが必要です。');
    if(Date.now()>sessionExpires-60000){
        const next=await request('/auth/v1/token?grant_type=refresh_token',{method:'POST',body:{refresh_token:session.refresh_token}});
        setSession(next);
    }
    return session.access_token;
}
const rest=(table,query='',options={})=>request('/rest/v1/'+table+'?'+query,options);
async function getPublication(){
    const rows=await rest('baseball_publication','id=eq.1&select=snapshot_id,chunk_count,published_at');
    if(!Array.isArray(rows)||rows.length>1)throw new Error('公開情報の形式が不正です。');
    return rows[0]||null;
}
async function packSnapshot(snapshot){
    const source=new Blob([JSON.stringify(snapshot)]).stream();
    const bytes=new Uint8Array(await new Response(source.pipeThrough(new CompressionStream('gzip'))).arrayBuffer());
    let binary='';for(let i=0;i<bytes.length;i+=16384)binary+=String.fromCharCode(...bytes.subarray(i,i+16384));
    return 'gzip64:'+btoa(binary);
}
async function unpackSnapshot(text){
    if(!text.startsWith('gzip64:'))throw new Error('保存形式が不明です。');
    const bytes=Uint8Array.from(atob(text.slice(7)),c=>c.charCodeAt(0));
    const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
    const snapshot=JSON.parse(await new Response(stream).text());
    if(snapshot?.version!==1||!snapshot.data||!Array.isArray(snapshot.anchors)||Object.keys(labels).some(k=>!Array.isArray(snapshot.data[k])))throw new Error('保存データの形式が不正です。');
    return snapshot;
}
async function readCloud(){
    // 読み込み中に更新されても、異なる版の断片を混ぜない。
    for(let attempt=0;attempt<3;attempt++){
        const publication=await getPublication();
        if(!publication)return {snapshot:null,version:null};
        const id=publication.snapshot_id,n=publication.chunk_count;
        if(!/^[0-9a-f-]{36}$/i.test(id)||!Number.isInteger(n)||n<1||n>10000)throw new Error('公開情報が不正です。');
        const pieces=[];
        for(let i=0;i<n;i+=4){
            const rows=await rest('baseball_snapshot_chunks',`snapshot_id=eq.${id}&part_no=gte.${i}&part_no=lt.${Math.min(i+4,n)}&select=part_no,payload&order=part_no.asc`);
            for(const row of rows)pieces.push(row);
        }
        const latest=await getPublication();
        if(latest?.snapshot_id!==id)continue;
        if(pieces.length!==n||pieces.some((p,i)=>p.part_no!==i||typeof p.payload!=='string'))throw new Error('公開データが不足しています。');
        const snapshot=await unpackSnapshot(pieces.map(p=>p.payload).join(''));
        return {snapshot,version:id};
    }
    throw new Error('データ更新中です。ページを開き直してください。');
}
async function publishCloud(snapshot){
    const previous=await getPublication();
    if((previous?.snapshot_id||null)!==cloudVersion)throw new Error('別の画面で更新されています。ページを開き直してから更新してください。');
    const encoded=await packSnapshot(snapshot),id=crypto.randomUUID(),parts=[];
    for(let i=0;i<encoded.length;i+=400000)parts.push(encoded.slice(i,i+400000));
    for(let i=0;i<parts.length;i++){
        publishStatus.textContent=`クラウドへ保存中：${i+1} / ${parts.length}`;
        await rest('baseball_snapshot_chunks','',{method:'POST',token:await adminToken(),body:{snapshot_id:id,part_no:i,payload:parts[i]},prefer:'return=minimal'});
    }
    const body={id:1,snapshot_id:id,chunk_count:parts.length,published_at:snapshot.publishedAt};
    const options={method:previous?'PATCH':'POST',token:await adminToken(),body,prefer:'return=representation'};
    const query=previous?`id=eq.1&snapshot_id=eq.${previous.snapshot_id}`:'';
    try{
        const rows=await rest('baseball_publication',query,options);
        if(rows?.length!==1||rows[0].snapshot_id!==id)throw new Error('公開版の切り替えに失敗しました。');
    }catch(e){
        // 応答だけが途切れた場合は公開先を読み直して確認。
        const current=await getPublication();if(current?.snapshot_id!==id)throw e;
    }
    cloudVersion=id;
    // 古い版はここでは削除しない。途中失敗時にも既存データを残す。
}
const localOnly=location.hostname==='localhost'||location.hostname==='127.0.0.1'||location.hostname==='[::1]';
const viewer=new URL(location.href);viewer.search='';viewer.hash='';const viewerURL=viewer.href;
const publishBox=document.createElement('section');publishBox.style.cssText=panel.style.cssText;
const heading=document.createElement('h3');heading.textContent='更新管理（クラウド）';publishBox.append(heading);
const note=document.createElement('p');note.style.cssText='line-height:1.6;margin-bottom:8px';note.textContent='Excel（5シート）・従来CSV・リーグ戦動画登録・途中補正は、ブラウザ内へ一時保存後、最後に「更新終了・一括保存」を押すと一括公開されます。';publishBox.append(note);
const publishStatus=document.createElement('p');publishStatus.setAttribute('aria-live','polite');publishStatus.style.cssText='line-height:1.6;overflow-wrap:anywhere;margin:6px 0';publishBox.append(publishStatus);
const publishButton=document.createElement('button');publishButton.type='button';publishButton.className='search-button';publishButton.textContent='更新終了・一括保存';publishButton.disabled=true;
const linkBox=document.createElement('div');linkBox.hidden=true;linkBox.style.cssText='margin-top:12px;overflow-wrap:anywhere;line-height:1.6';
const link=document.createElement('a');link.href=viewerURL;link.target='_blank';link.rel='noopener';link.textContent=localOnly?'閲覧者ページを開く（ローカル確認用）':'閲覧者ページを開く';linkBox.append(link);
const urlText=document.createElement('input');urlText.readOnly=true;urlText.value=viewerURL;urlText.style.cssText='width:100%;margin-top:8px;padding:6px;border:1px solid #b9c5d3;border-radius:4px';linkBox.append(urlText);
const copy=document.createElement('button');copy.type='button';copy.className='reset-button';copy.textContent='閲覧用URLをコピー';copy.addEventListener('click',async()=>{try{await navigator.clipboard.writeText(viewerURL);copy.textContent='コピーしました';}catch{urlText.focus();urlText.select();}});linkBox.append(copy);
const viewerError=document.createElement('p');viewerError.setAttribute('role','alert');viewerError.style.cssText='color:#b42318;font-size:12px;margin-top:12px';viewerError.hidden=true;
if(!adminMode)form.after(viewerError);
function showLoadError(message){if(adminMode)publishStatus.textContent=message;else{viewerError.textContent=message;viewerError.hidden=false;}}
let lastPublished=null;
const totalRows=()=>Object.values(dataByCategory).reduce((n,rows)=>n+rows.length,0);
function updatePublishUI(){
    if(!adminMode)return;
    publishButton.disabled=!session||!ready||busy>0||publishing||!dirty||leagueDraftPending()||totalRows()===0;
    if(ready&&!publishing)publishStatus.textContent=`編集中：${totalRows().toLocaleString('ja-JP')}件｜${leagueDraftPending()?'リーグ戦の未保存入力あり（先に一時保存）':dirty?'一時保存済み・未公開':'変更なし'}${lastPublished?'｜最終確定：'+new Date(lastPublished).toLocaleString('ja-JP'):''}`;
}
function applySnapshot(snapshot){
    if(snapshot&&snapshot.version!==1)throw new Error('保存データの形式が異なります。');
    for(const category of Object.keys(labels)){
        dataByCategory[category]=snapshot?.data?.[category]||[];
        const status=uploadStatuses.get(category);
        status.textContent=dataByCategory[category].length?`${dataByCategory[category].length.toLocaleString('ja-JP')}件 読み込み済み`:'未読み込み';
        status.style.color=dataByCategory[category].length?'#17633a':'#64748b';
    }
    leagueAnchors.clear();invalidateVideoSearchCache();for(const [key,list] of snapshot?.anchors||[])leagueAnchors.set(key,list);
    lastPublished=snapshot?.publishedAt||null;dirty=false;
    refreshLeaguePanel();populateFilters();search();
    linkBox.hidden=!cloudVersion;
}
async function loadViewer(){
    try{const loaded=await readCloud();cloudVersion=loaded.version;applySnapshot(loaded.snapshot);ready=true;}
    catch(e){showLoadError('クラウド読み込み失敗：'+e.message+'。ページを開き直してください。');}
}
let logoutControl=null;
async function commitPending(){
    if(!session||!ready||busy||publishing||!dirty||leagueDraftPending()||!totalRows())return false;
    publishing=true;publishButton.disabled=true;if(logoutControl)logoutControl.disabled=true;panel.inert=true;const leagueBox=panel.nextElementSibling;if(leagueBox)leagueBox.inert=true;publishStatus.textContent='保存中…';
    try{
        const snapshot={version:1,publishedAt:new Date().toISOString(),data:dataByCategory,anchors:[...leagueAnchors]};
        await persistDraft();
        await savePublished(snapshot,'migration-backup');
        await publishCloud(snapshot);
        try{await savePublished(snapshot);}catch{}
        lastPublished=snapshot.publishedAt;dirty=false;linkBox.hidden=false;
        let cleanupWarning='';try{await savePublished(null,draftKey);}catch{cleanupWarning='（公開成功・下書きの削除のみ失敗）';}
        publishStatus.textContent=`クラウド更新完了：${totalRows().toLocaleString('ja-JP')}件`+cleanupWarning;
        return true;
    }catch(e){publishStatus.textContent='更新失敗または結果未確認：'+e.message+'。公開済みデータを確認して再試行してください。';return false;}
    finally{publishing=false;panel.inert=false;if(leagueBox)leagueBox.inert=false;if(logoutControl)logoutControl.disabled=false;publishButton.disabled=!session||!ready||busy>0||!dirty||leagueDraftPending()||!totalRows();}
}
if(adminMode){
    panel.hidden=true;for(const leaguePanel of leaguePanels)leaguePanel.hidden=true;publishBox.hidden=true;
    const login=document.createElement('form');login.style.cssText=panel.style.cssText;
    const title=document.createElement('h3');title.textContent='管理者ログイン';login.append(title);
    const makeInput=(name,type)=>{const label=document.createElement('label');label.textContent=name;label.style.cssText='display:block;margin:8px 0';const el=document.createElement('input');el.type=type;el.required=true;el.autocomplete=type==='password'?'current-password':'username';el.style.cssText='display:block;width:100%;padding:8px;border:1px solid #b9c5d3;border-radius:4px';label.append(el);login.append(label);return el;};
    const email=makeInput('メールアドレス','email'),password=makeInput('管理者用パスワード','password');
    const loginButton=document.createElement('button');loginButton.type='submit';loginButton.className='search-button';loginButton.textContent='ログイン';login.append(loginButton);
    const loginStatus=document.createElement('p');loginStatus.setAttribute('role','status');loginStatus.style.cssText='margin-top:8px;overflow-wrap:anywhere';login.append(loginStatus);
    panel.before(login);form.before(publishBox);publishBox.append(publishButton,linkBox);
    const logout=document.createElement('button');logout.type='button';logout.className='reset-button';logout.textContent='ログアウト';publishBox.append(logout);logoutControl=logout;
    logout.addEventListener('click',async()=>{if(publishing||busy)return;if(leagueDraftPending()&&!confirm('リーグ戦の未保存入力だけが失われます。一時保存済みデータは残ります。ログアウトしますか？'))return;try{await request('/auth/v1/logout?scope=local',{method:'POST',token:await adminToken()});}catch{}session=null;dirty=false;location.reload();});
    login.addEventListener('submit',async e=>{
        e.preventDefault();if(loginButton.disabled)return;loginButton.disabled=true;loginStatus.textContent='ログイン中…';
        try{
            const value=await request('/auth/v1/token?grant_type=password',{method:'POST',body:{email:email.value,password:password.value}});password.value='';setSession(value);
            const loaded=await readCloud();cloudVersion=loaded.version;
            let snapshot=loaded.snapshot,migration=false;
            if(!snapshot&&localOnly){snapshot=await readPublished();migration=!!snapshot;}
            const draft=await readPublished(draftKey);
            let restored=false;
            if(draft){
                if(draft.schema!==1||draft.baseVersion!==cloudVersion){
                    if(!confirm('別の公開版に基づく下書きがあります。最新のクラウドデータを上書きしないため復元できません。古い下書きを残したままログインを中止する場合はキャンセル、破棄する場合はOKを押してください。'))throw new Error('古い下書きを保持しました。');
                    await savePublished(null,draftKey);
                }else{snapshot=draft.snapshot;resumeMatch=draft.nextMatch||'';restored=true;}
            }
            applySnapshot(snapshot);ready=true;dirty=migration||restored;if(restored&&resumeMatch)refreshLeaguePanel(resumeMatch);
            panel.hidden=false;for(const leaguePanel of leaguePanels)leaguePanel.hidden=leaguePanel.dataset.view!==activeView;publishBox.hidden=false;login.hidden=true;
            updatePublishUI();
            if(restored)publishStatus.textContent=`一時保存データを復元：${totalRows().toLocaleString('ja-JP')}件｜未公開。最後に更新終了を押してください。`;
            else if(migration)publishStatus.textContent=`ローカル確定版 ${totalRows().toLocaleString('ja-JP')}件を復元。公開許可を確認後、更新終了を押すと初回クラウド移行します。`;
        }catch(err){session=null;ready=false;password.value='';loginStatus.textContent='ログイン・読み込み失敗：'+err.message;}
        finally{loginButton.disabled=false;}
    });
    publishButton.addEventListener('click',async()=>{
        if(publishButton.disabled)return;
        if(!confirm('ログイン不要で閲覧できるデータとして公開します。部内の許可を得た選手情報・動画・シートBT所属を確認しましたか？'))return;
        await commitPending();
    });
    window.addEventListener('beforeunload',e=>{if(busy||publishing||leagueDraftPending()){e.preventDefault();e.returnValue='';}});
    if(!configured())loginStatus.textContent='まず script.js の公開用キーを設定してください。';
}else loadViewer();


// STEP 36: 初回アクセスはサイト選択画面。既存URLは維持し、#batter / #pitcher で直接入場可能。
(function installEntrance(){
    const styles=document.createElement('style');styles.textContent=`
    .bvs-landing{position:fixed;inset:0;z-index:10000;background:linear-gradient(145deg,#08182d,#193c63);color:#fff;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box;overflow:auto}
    .bvs-landing[hidden]{display:none!important}
    .bvs-landing-inner{max-width:880px;width:100%;text-align:center}
    .bvs-landing h1{color:#fff;font-size:clamp(25px,4vw,42px);margin:0 0 36px;line-height:1.3}
    .bvs-landing-cards{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}
    .bvs-landing-card{border:1px solid #6e8aab;background:#fff;color:#132b46;border-radius:14px;padding:32px 18px;cursor:pointer;min-height:190px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;box-shadow:0 8px 25px #07101c55}
    .bvs-landing-card:hover,.bvs-landing-card:focus-visible{outline:3px solid #9dccff;outline-offset:2px;transform:translateY(-2px)}
    .bvs-landing-card .bvs-icon{display:inline-flex;width:84px;height:84px;align-items:center;justify-content:center;flex:none}
    .bvs-landing-card .bvs-icon svg{display:block;width:100%;height:100%}
    .bvs-landing-card .bvs-icon img{display:block;width:100%;height:100%;object-fit:contain}
    .bvs-landing-card strong{font-size:23px}
    .bvs-switcher{position:sticky;top:0;z-index:950;background:#102b49;color:#fff;padding:9px 14px;display:flex;align-items:center;gap:10px;justify-content:space-between;box-sizing:border-box}
    .bvs-switcher[hidden]{display:none!important}
    .bvs-switcher strong{font-size:14px;color:#fff}
    .bvs-switcher button{border:1px solid #b5d3ee;border-radius:5px;background:#fff;color:#174575;padding:6px 12px;font-weight:700;cursor:pointer}
    @media(max-width:620px){.bvs-landing-cards{grid-template-columns:1fr}.bvs-landing-card{min-height:130px;padding:20px}}
    `;document.head.append(styles);
    const landing=document.createElement('section');landing.className='bvs-landing';landing.setAttribute('aria-label','野手・投手動画検索の選択');
    const inner=document.createElement('div');inner.className='bvs-landing-inner';
    const title=document.createElement('h1');title.textContent='野球動画検索システム';
    const cards=document.createElement('div');cards.className='bvs-landing-cards';
    // バットSVGと野球ボール写真はコード内に埋め込み。GitHubにはscript.jsだけをアップロードできます。
    const batSvg=`<svg viewBox="0 0 96 96" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
      <defs><linearGradient id="bvs-bat-wood37" x1="0" y1="0" x2="1" y2="0"><stop stop-color="#b87734"/><stop offset=".45" stop-color="#e6af66"/><stop offset="1" stop-color="#925326"/></linearGradient></defs>
      <g transform="rotate(38 48 48)">
        <path d="M38 14 Q38 8 44 8 H52 Q58 8 58 14 V33 Q58 45 53 54 L53 79 H43 L43 54 Q38 45 38 33 Z" fill="url(#bvs-bat-wood37)" stroke="#895023" stroke-width="1.8" stroke-linejoin="round"/>
        <path d="M42 14 V34 Q42 43 46 51" fill="none" stroke="#f8cf94" stroke-width="1.8" opacity=".75" stroke-linecap="round"/>
        <path d="M45 63 H51 M45 67 H51 M45 71 H51 M45 75 H51" stroke="#63412c" stroke-width="2" stroke-linecap="round"/>
        <path d="M43 78 H53 V84 H43 Z" fill="#70412c"/><rect x="40" y="84" width="16" height="5" rx="2.5" fill="#a86a38" stroke="#754526" stroke-width="1"/>
      </g>
    </svg>`;
    // STEP 38: ユーザー提供の野球ボール写真をWebPにしてコード内に埋め込み。画像ファイルの別アップロード不要。
    const ballPhoto='<img src="data:image/webp;base64,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" alt="" loading="eager">';
    function card(view,svg,title){
        const b=document.createElement('button');b.type='button';b.className='bvs-landing-card';
        const figure=document.createElement('span');figure.className='bvs-icon';figure.setAttribute('aria-hidden','true');figure.innerHTML=svg;
        const t=document.createElement('strong');t.textContent=title;
        b.append(figure,t);b.addEventListener('click',()=>enter(view));cards.append(b);
    }
    card('batter',batSvg,'野手動画');
    card('pitcher',ballPhoto,'投手動画');
    inner.append(title,cards);landing.append(inner);
    const switcher=document.createElement('div');switcher.className='bvs-switcher';switcher.hidden=true;
    const current=document.createElement('strong');
    const back=document.createElement('button');back.type='button';back.textContent='◀ 動画の種類を選び直す';
    switcher.append(current,back);document.body.prepend(switcher,landing);
    function choose(){landing.hidden=false;switcher.hidden=true;document.body.style.overflow='hidden';location.hash='';}
    back.addEventListener('click',choose);
    function enter(view){
        if(!['batter','pitcher'].includes(view))return;
        activeView=view;
        for(const [key] of filterFields){filterSelections[key].clear();filterPanels[key].query.value='';filterPanels[key].details.open=false;}
        filterPanels.pitchType.group.style.display=view==='pitcher'?'':'none';
        for(const box of leaguePanels)box.hidden=!adminMode||box.dataset.view!==view||panel.hidden;
        current.textContent=(view==='pitcher'?'投手動画':'野手動画')+' 検索';
        landing.hidden=true;switcher.hidden=false;document.body.style.overflow='';
        location.hash=view;
        populateFilters();search();
    }
    const hash=location.hash.replace('#','');
    if(hash==='batter'||hash==='pitcher')enter(hash);else choose();
})();