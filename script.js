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
function markDirty(){dirty=true;updatePublishUI();}

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
        if(!batter||!dateMatch||(!sheet&&get('打者チーム')!=='慶應')) continue;
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
        normalized.push({sourceRow,pitchInAtBat:get('打席内球数'),category, date, year:dateMatch[1], opponent:sheet?'':get('投手チーム'), batter, pitcher:get(sheet?'投手':'投手名'), batterSide:sheet?'':side(get('打者打席左右')), pitcherSide:sheet?'':side(get('投手投げ手左右')), inning:sheet||!get('イニング')?'':get('イニング')+'回'+get('表/裏'), count:sheet||balls===''||strikes===''?'':balls+'-'+strikes, pitchType:measurementValid?(typeMap[rawType]||rawType):'', speed, pitchResult, atBatResult, result:atBatResult||pitchResult, resultClass:classifyResult(atBatResult||pitchResult), timestamp:sheet?get('Time'):get('タイムスタンプ'), angles});
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
const form=$('searchForm');
// 既存HTMLを変更せず検索フォームだけを再構築する。
const filterFields = [['year','年度'],['date','日付'],['category','試合区分'],['opponent','対戦大学'],['pitcher','投手'],['batter','打者'],['result','打撃結果']];
// 保存済みデータの古い分類を参照せず、最新の5分類で都度判定する。
const filterValue = (record,key) => key==='result' ? classifyResult(record.result || record.atBatResult || record.pitchResult) : record[key];
form.replaceChildren();
for(const [key,name] of filterFields){
    const group=document.createElement('div');group.className='filter-group';
    const label=document.createElement('label');label.htmlFor=key;label.textContent=name;
    const select=document.createElement('select');select.id=key;
    select.append(new Option('すべて',''));
    group.append(label,select);form.append(group);
}
const searchButton=document.createElement('button');searchButton.type='button';searchButton.id='searchButton';searchButton.className='search-button';searchButton.textContent='検索する';
const resetButton=document.createElement('button');resetButton.type='button';resetButton.id='clearFilters';resetButton.className='reset-button';resetButton.textContent='絞り込みを全解除';
form.append(searchButton,resetButton);
// テスト用CSV選択欄を追加。
const panel=document.createElement('details');
panel.open=true;
panel.style.cssText='margin-bottom:20px;padding:12px;background:#fff;border:1px solid #b9c5d3;border-radius:6px;font-size:12px;';
const summary=document.createElement('summary');summary.textContent='打者動画データ読み込み（Excel 1ファイル対応）';summary.style.cssText='font-weight:bold;cursor:pointer;margin-bottom:8px;';panel.append(summary);
const excelHint=document.createElement('p');excelHint.textContent='5つのシート（リーグ戦／OP戦(A)／OP戦(B)／紅白戦／シートBT）が入ったExcelを1つ選択してください。すべて読み取ってブラウザ内に一時保存し、「更新終了・一括保存」で初めて閲覧者側へ反映します。';excelHint.style.cssText='line-height:1.6;margin:8px 0 10px';panel.append(excelHint);
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
            if(remap){leagueAnchors.clear();for(const [k,v] of remap.entries)leagueAnchors.set(k,v);}
            await persistDraft();
            if(remap)refreshLeaguePanel();
            markDirty();
            status.textContent=`${records.length.toLocaleString('ja-JP')}件 読み込み完了`+(category==='sheet'?'（所属未確認）':'')+(remap?`｜リーグ戦基準点：${remap.kept}件引継ぎ、${remap.dropped}件要再登録`:'');
            status.style.color='#17633a';
            populateFilters();search();
        } catch(e) {dataByCategory[category]=oldRows;leagueAnchors.clear();for(const [k,v] of oldAnchors)leagueAnchors.set(k,v);dirty=oldDirty;status.textContent='読み込み失敗：'+e.message+'（以前のデータは保持）';status.style.color='#b42318';}
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
        excelStatus.textContent='Excelの5シートを解析しています…';
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
        leagueAnchors.clear();for(const [k,v] of remap.entries)leagueAnchors.set(k,v);
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
        leagueAnchors.clear();for(const [k,v] of oldAnchors)leagueAnchors.set(k,v);
        dirty=oldDirty;resumeMatch=oldResume;
        excelStatus.textContent='Excel読み込み失敗：'+e.message+'\n既存のデータ・動画設定は保持されています。';excelStatus.style.color='#b42318';
    }finally{
        excelInput.disabled=false;excelInput.value='';busy--;panel.inert=false;
        if(leagueBox)leagueBox.inert=false;updatePublishUI();
    }
});
if(adminMode)form.before(panel);
// 他の6項目の条件に合うデータから、各項目の選択肢を作成する。
function selectedFilters() {
    return Object.fromEntries(filterFields.map(([key])=>[key,$(key).value]));
}
function matchesFilters(record, selected, exceptKey=null) {
    return filterFields.every(([key])=>key===exceptKey||!selected[key]||filterValue(record,key)===selected[key]);
}
function availableFilterValues(all, selected, key) {
    const values=new Set();
    for(const record of all) {
        if(!matchesFilters(record,selected,key))continue;
        const value=filterValue(record,key);
        if(value)values.add(value);
    }
    if(key==='result')return Object.keys(resultGroups).filter(value=>values.has(value));
    if(key==='category')return Object.keys(labels).filter(value=>values.has(value));
    return [...values].sort((a,b)=>a.localeCompare(b,'ja',{numeric:true}));
}
function populateFilters() {
    const all=Object.values(dataByCategory).flat();
    let selected=selectedFilters();
    // CSV差し替えで組み合わせが消えた場合だけ、検索条件を解除する。
    if(Object.values(selected).some(Boolean)&&!all.some(r=>matchesFilters(r,selected))) {
        selected=Object.fromEntries(filterFields.map(([key])=>[key,'']));
    }
    // 同じ選択状態を基準に、7項目すべてを更新する（更新順に依存しない）。
    for(const [key] of filterFields) {
        const values=availableFilterValues(all,selected,key), sel=$(key);
        sel.replaceChildren(new Option('すべて',''),...values.map(v=>new Option(key==='category'?labels[v]:v,v)));
        sel.value=values.includes(selected[key])?selected[key]:'';
    }
}
function search() {
    const selected=selectedFilters();
    results=Object.values(dataByCategory).flat().filter(r=>matchesFilters(r,selected));
    page=1;clearSelection();renderTable();
}
function renderTable() {
    const size=Number($('pageSize').value)||50,total=Math.ceil(results.length/size);
    page=Math.max(1,Math.min(page,total||1));
    $('resultCount').textContent=results.length.toLocaleString('ja-JP')+'件';
    $('pageInfo').textContent=`${total?page:0} / ${total} ページ`;
    $('prevPage').disabled=page<=1;$('nextPage').disabled=page>=total;
    const body=$('resultBody');body.replaceChildren();
    if(!results.length){const row=body.insertRow(),cell=row.insertCell();cell.colSpan=12;cell.className='empty-table';cell.textContent='該当するデータがありません';return;}
    const start=(page-1)*size;
    for(const [i,r] of results.slice(start,start+size).entries()) {
        const tr=body.insertRow();
        const index=start+i;
        tr.classList.toggle("selected",index===selectedIndex);
        const values=[start+i+1,labels[r.category],r.date,r.opponent,r.batter,r.pitcher,r.inning,r.count,r.pitchType,r.speed===null?'':r.speed+' km/h',r.result];
        for(const value of values){const td=tr.insertCell();td.textContent=value===''?'―':String(value);}
        const cell=tr.insertCell(), button=document.createElement('button');
        const playable=Object.keys(playableAngles(r)).length>0;
        button.type='button';button.className='play-button';button.textContent=playable?'▶ 再生':'詳細';
        button.title=playable?'該当する動画を再生':'動画の再生位置が未登録です';
        button.addEventListener('click',()=>selectVideo(index,true));cell.append(button);
    }
}
$('searchButton').addEventListener('click',search);
form.addEventListener('submit',e=>{e.preventDefault();search();});
// 選択肢だけを即時更新し、動画・一覧は検索ボタンまで変更しない。
for(const [key] of filterFields)$(key).addEventListener('change',populateFilters);
function clearAllFilters(){
    for(const [key] of filterFields)$(key).value='';
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
function playableAngles(record) {
    if(record.category==='league') return Object.fromEntries(angleOrder.map(a=>[a,leagueAngle(record,a)]).filter(([,v])=>v!==null));
    return Object.fromEntries(Object.entries(record.angles).map(([key,val])=>[key,parseYouTube(val)]).filter(([,val])=>val!==null));
}
const angleOrder=['ネット裏','センター裏','一塁側','三塁側','一塁側(投手)','三塁側(投手)'];
const angleButtons=[...document.querySelectorAll('.angle-button')];
angleButtons.forEach((button,i)=>{if(angleOrder[i]==='センター裏')button.textContent=angleLabel(angleOrder[i]);});
const player=$('videoPlayer');
const videoLink=document.createElement('a');
videoLink.textContent='YouTubeで開く ↗';videoLink.target='_blank';videoLink.rel='noopener';videoLink.hidden=true;
videoLink.style.cssText='display:none;color:#24558e;font-size:13px;margin-top:10px;text-align:right;';
player.after(videoLink);
const showInfo=(r)=>{
    const values={infoBatter:r?.batter,infoPitcher:r?.pitcher,infoPitchType:r?.pitchType,infoSpeed:r?.speed!=null?r.speed+' km/h':'',infoCount:r?.count,infoResult:r?.result};
    for(const [id,value] of Object.entries(values))$(id).textContent=value||'―';
};
function placeholder(message){
    const div=document.createElement('div');div.className='video-placeholder';div.style.padding='16px';
    const icon=document.createElement('div');icon.className='play-icon';icon.textContent='▶';
    const p=document.createElement('p');p.textContent=message;div.append(icon,p);player.replaceChildren(div);
    videoLink.hidden=true;videoLink.style.display='none';videoLink.removeAttribute('href');
}
function clearSelection(){
    selectedIndex=-1;selectedAngle=null;showInfo(null);placeholder('動画を選択してください');
    $('prevVideo').disabled=true;$('nextVideo').disabled=true;$('videoPosition').textContent='0 / '+results.length;
    angleButtons.forEach(b=>{b.disabled=true;b.classList.remove('active');});
}
function displayAngle(key){
    const r=results[selectedIndex];if(!r)return;
    const angles=playableAngles(r), video=angles[key];if(!video)return;
    selectedAngle=key;
    angleButtons.forEach((b,i)=>{b.disabled=!angles[angleOrder[i]];b.classList.toggle('active',angleOrder[i]===key);});
    const watch=new URL('https://www.youtube.com/watch');watch.searchParams.set('v',video.id);watch.searchParams.set('t',video.start+'s');
    videoLink.href=watch.href;videoLink.hidden=false;videoLink.style.display='block';
    if(location.protocol==='file:') {
        // file:// ではReferer不足によるエラーが起こりうるためリンクを案内。
        placeholder('サイト内再生にはローカルサーバーで開いてください。下のリンクからYouTubeで再生できます。');
        videoLink.href=watch.href;videoLink.hidden=false;videoLink.style.display='block';return;
    }
    const src=new URL('https://www.youtube.com/embed/'+video.id);
    src.searchParams.set('start',String(video.start));src.searchParams.set('autoplay','1');src.searchParams.set('playsinline','1');
    const iframe=document.createElement('iframe');iframe.title=r.batter+'｜'+angleLabel(key)+'｜打撃動画';
    iframe.allow='accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share';
    iframe.allowFullscreen=true;iframe.referrerPolicy='strict-origin-when-cross-origin';iframe.src=src.href;
    player.replaceChildren(iframe);
}
function selectVideo(index,scroll=false){
    if(index<0||index>=results.length)return;
    selectedIndex=index;const r=results[index], angles=playableAngles(r);
    showInfo(r);$('videoPosition').textContent=`${index+1} / ${results.length}`;
    $('prevVideo').disabled=index===0;$('nextVideo').disabled=index===results.length-1;
    const preferred=angles[selectedAngle]?selectedAngle:angleOrder.find(a=>angles[a]);
    if(preferred){displayAngle(preferred);}else{
        placeholder(r.category==='league'?'リーグ戦：動画の基準時刻が未登録、または撮影中断区間です。':'この投球の動画URLまたは再生時刻が未登録です。');
        angleButtons.forEach(b=>{b.disabled=true;b.classList.remove('active');});
    }
    page=Math.floor(index/(Number($('pageSize').value)||50))+1;renderTable();
    if(scroll)player.scrollIntoView({behavior:'smooth',block:'start'});
}
angleButtons.forEach((b,i)=>b.addEventListener('click',()=>displayAngle(angleOrder[i])));
$('prevVideo').addEventListener('click',()=>selectVideo(selectedIndex-1,true));
$('nextVideo').addEventListener('click',()=>selectVideo(selectedIndex+1,true));
clearSelection();

// STEP 8/30: リーグ戦の同期。登録はブラウザの一時保存に保持。
const leagueAnchors = new Map();
const leagueMatchKey = r => JSON.stringify([r.date,r.opponent]);
const leagueGroupKey = r => JSON.stringify([r.date,r.opponent,r.inning]);
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
function setupLeaguePanel() {
    const box=document.createElement('details');box.style.cssText=panel.style.cssText;
    const summary=document.createElement('summary');summary.textContent='リーグ戦 動画登録（1試合ずつ）';summary.style.cssText='font-weight:bold;cursor:pointer;margin-bottom:8px';box.append(summary);
    const p=(parent,t)=>{const el=document.createElement('p');el.textContent=t;el.style.cssText='margin:8px 0;line-height:1.5;overflow-wrap:anywhere;';parent.append(el);return el;};
    p(box,'試合を選択し、各アングルの元チャプター・投球動作開始2秒前の修正時刻・YouTube URLを登録してください。再生は修正時刻のさらに3秒前から。片方の映像しかない場合はもう片方の3項目を空欄のままにしてください。各試合を一時保存した後、最後に更新管理欄の「更新終了」を押すとまとめて公開されます。');
    function field(parent,label,el){const l=document.createElement('label');l.textContent=label;l.style.cssText='display:block;font-weight:bold;margin:10px 0 4px';el.style.cssText='display:block;width:100%;margin-top:4px;padding:6px;border:1px solid #b9c5d3;border-radius:4px;font-size:12px;box-sizing:border-box;';l.append(el);parent.append(l);return el;}
    const match=field(box,'試合',document.createElement('select'));
    const fields=new Map();let edited=false, currentMatch='';leagueDraftPending=()=>edited;
    const progress=p(box,'');
    for(const a of leagueRegistrationAngles){
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
    function updateProgress(){const keys=new Set(dataByCategory.league.map(leagueMatchKey)), registered=new Set(dataByCategory.league.filter(r=>leagueRegistrationAngles.some(a=>leagueAngle(r,a))).map(leagueMatchKey));progress.textContent='再生位置登録済み（片方以上）：'+registered.size+' / '+keys.size+'試合｜未公開の一時保存は最後に更新終了';}
    const rows=()=>dataByCategory.league.filter(r=>leagueMatchKey(r)===match.value);
    const requireMatch=()=>{if(!rows().length)throw new Error('リーグ戦CSVと試合を選択してください。');};
    const count=()=>rows().filter(r=>leagueRegistrationAngles.some(a=>leagueAngle(r,a))).length;
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
        const previous=preferred;const m=new Map(dataByCategory.league.map(r=>[leagueMatchKey(r),r.date+' vs '+r.opponent]));
        match.replaceChildren(new Option('選択してください',''),...[...m].sort((x,y)=>x[1].localeCompare(y[1])).map(([k,label])=>new Option(label,k)));
        if(m.has(previous))match.value=previous;
        populateInputs();say(match.value?'この試合の再生位置登録済み：'+count()+' / '+rows().length+'件':'試合を選択してください。');
    }
    match.addEventListener('change',()=>{
        if(busy||publishing){match.value=currentMatch;return;}
        if(edited&&!confirm('未保存の6項目を破棄して別の試合を選択しますか？')){match.value=currentMatch;return;}
        populateInputs();updatePublishUI();say('この試合の再生位置登録済み：'+count()+' / '+rows().length+'件');
    });
    function move(delta){if(!ready||busy||publishing)return;const target=neighbor(delta);if(!target)return;if(edited&&!confirm('未保存の6項目を破棄して移動しますか？'))return;match.value=target;populateInputs();updatePublishUI();say('この試合の再生位置登録済み：'+count()+' / '+rows().length+'件');}
    prev.addEventListener('click',()=>move(-1));skip.addEventListener('click',()=>move(1));
    save.addEventListener('click',async()=>{
        if(!ready||!session||publishing||busy)return;
        try{
            requireMatch();
            const plan=[...fields].flatMap(([a,f])=>JSON.stringify([f.url.value,f.raw.value,f.corrected.value])===f.initial?[]:planLeagueAngle(rows(),a,{url:f.url.value,raw:f.raw.value,corrected:f.corrected.value}));
            if(!plan.length)throw new Error('保存する変更がありません。少なくとも片方のアングルの3項目を入力・変更してください。');
            const label=match.selectedOptions[0].textContent;
            const before=new Map(leagueAnchors),oldDirty=dirty,oldResume=resumeMatch;
            const next=neighbor(1);
            busy++;box.inert=true;updatePublishUI();
            try{
                for(const {r,angle,anchor} of plan)putAnchor(r,angle,anchor);
                resumeMatch=next||match.value;
                await persistDraft();
            }catch(e){leagueAnchors.clear();for(const [k,v] of before)leagueAnchors.set(k,v);dirty=oldDirty;resumeMatch=oldResume;say('一時保存失敗：'+e.message+'。入力は保持しています。',true);return;}
            finally{busy--;box.inert=false;updatePublishUI();}
            markDirty();const savedCount=count(),total=rows().length;
            if(next)match.value=next;
            populateInputs();search();updatePublishUI();
            say(`${label}：一時保存完了（再生位置 ${savedCount} / ${total}件）。`+(next?' 次の試合を入力できます。':' 最後の試合です。更新終了で公開してください。'));
        }catch(e){say(e.message,true);}
    });
    // 例外対応は折りたたみ内に残す。通常登録は上の6入力のみ。
    const advanced=document.createElement('details');advanced.style.cssText='margin-top:16px;border-top:1px solid #cdd7e3;padding-top:10px;';box.append(advanced);
    const as=document.createElement('summary');as.textContent='途中補正・撮影中断（必要な場合のみ）';advanced.append(as);
    p(advanced,'上の6項目に未保存入力がある場合は先に試合を保存してください。中断した最初の投球に「動画なし」、再開した最初の投球に新しい基準時刻を登録。補正も一時保存され、下の更新終了でまとめて公開します。');
    const angle=field(advanced,'撮影アングル',document.createElement('select'));angle.replaceChildren(...leagueRegistrationAngles.map(a=>new Option(angleLabel(a),a)));
    const url=field(advanced,'再開動画のYouTube URL（空欄なら上のURL）',document.createElement('input'));url.type='url';
    const inning=field(advanced,'イニング',document.createElement('select'));
    const pitch=field(advanced,'基準にする投球（CSV順）',document.createElement('select'));
    const time=field(advanced,'動画内の時刻（投球動作の2秒前）',document.createElement('input'));time.placeholder='1:23:45';
    function updatePitches(){const previous=pitch.value;const a=rows().filter(r=>r.inning===inning.value);pitch.replaceChildren(new Option('選択してください',''),...a.map((r,i)=>{const item=(leagueAnchors.get(anchorSlot(r,angle.value))||[]).find(x=>x.row===r.sourceRow);return new Option((i+1)+'球目｜'+r.timestamp.slice(-8)+'｜'+r.batter+'｜打席内'+r.pitchInAtBat+'球'+(item?item.missing?'【動画なし】':'【基準登録済】':''),String(r.sourceRow));}));if([...pitch.options].some(o=>o.value===previous))pitch.value=previous;}
    function updateInnings(){const previous=inning.value;inning.replaceChildren(new Option('選択してください',''),...[...new Set(rows().map(r=>r.inning).filter(Boolean))].map(i=>new Option(i,i)));if([...inning.options].some(o=>o.value===previous))inning.value=previous;updatePitches();}
    function button(label,action){const b=document.createElement('button');b.type='button';b.textContent=label;b.className='search-button';b.style.cssText='margin:6px 0;font-size:12px;padding:8px;';b.addEventListener('click',async()=>{if(!ready||publishing||busy)return;if(edited){say('先に上の6項目を一時保存してください。',true);return;}const before=new Map(leagueAnchors),oldDirty=dirty,oldResume=resumeMatch;busy++;box.inert=true;updatePublishUI();try{action();resumeMatch=match.value;await persistDraft();markDirty();search();populateInputs();say('途中補正を一時保存しました。下の「更新終了」で公開してください。');}catch(e){leagueAnchors.clear();for(const [k,v] of before)leagueAnchors.set(k,v);dirty=oldDirty;resumeMatch=oldResume;say(e.message,true);}finally{busy--;box.inert=false;updatePublishUI();}});advanced.append(b);}
    const chosen=()=>{requireMatch();const r=rows().find(r=>r.inning===inning.value&&String(r.sourceRow)===pitch.value);if(!r)throw new Error('基準にする投球を選択してください。');return r;};
    inning.addEventListener('change',updatePitches);angle.addEventListener('change',updatePitches);
    button('この投球から再同期',()=>{const r=chosen(),v=parseYouTube({url:url.value||fields.get(angle.value).url.value,seconds:'0'}),start=parseTime(time.value);if(!v||start===null||clockSeconds(r.timestamp)===null)throw new Error('動画URL・動画内時刻・CSV時刻を確認してください。');putAnchor(r,angle.value,{id:v.id,start});});
    button('この投球から動画なし',()=>putAnchor(chosen(),angle.value,{missing:true}));
    button('選択した投球の基準・中断点を削除',()=>{const r=chosen(),key=anchorSlot(r,angle.value);leagueAnchors.set(key,(leagueAnchors.get(key)||[]).filter(a=>a.row!==r.sourceRow));});
    panel.after(box);updateMatches();return updateMatches;
}
const refreshLeaguePanel=adminMode?setupLeaguePanel():()=>{};
populateFilters();


// STEP 16: CSV差し替え時に行番号だけで同期点を引き継がない。
function remapLeagueAnchors(oldRows,newRows){
    const signature=r=>JSON.stringify([r.date,r.opponent,r.inning,r.timestamp,r.batter,r.pitcher,r.pitchInAtBat,r.count]);
    const index=rows=>{const m=new Map();for(const r of rows){const k=signature(r);if(m.has(k))m.set(k,null);else m.set(k,r);}return m;};
    const oldByRow=new Map(oldRows.map(r=>[r.sourceRow,r]));
    const oldIndex=index(oldRows),newIndex=index(newRows),entries=new Map();
    let kept=0,dropped=0;
    for(const [key,list] of leagueAnchors){
        const [group,angle]=JSON.parse(key);
        const pending=[];let failed=false;
        for(const a of list){
            const old=oldByRow.get(a.row),sig=old&&signature(old),next=sig&&newIndex.get(sig);
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
    if(ready&&!publishing)publishStatus.textContent=`編集中：${totalRows().toLocaleString('ja-JP')}件｜${leagueDraftPending()?'6項目の未保存入力あり（先に一時保存）':dirty?'一時保存済み・未公開':'変更なし'}${lastPublished?'｜最終確定：'+new Date(lastPublished).toLocaleString('ja-JP'):''}`;
}
function applySnapshot(snapshot){
    if(snapshot&&snapshot.version!==1)throw new Error('保存データの形式が異なります。');
    for(const category of Object.keys(labels)){
        dataByCategory[category]=snapshot?.data?.[category]||[];
        const status=uploadStatuses.get(category);
        status.textContent=dataByCategory[category].length?`${dataByCategory[category].length.toLocaleString('ja-JP')}件 読み込み済み`:'未読み込み';
        status.style.color=dataByCategory[category].length?'#17633a':'#64748b';
    }
    leagueAnchors.clear();for(const [key,list] of snapshot?.anchors||[])leagueAnchors.set(key,list);
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
    const leaguePanel=panel.nextElementSibling;
    panel.hidden=true;if(leaguePanel)leaguePanel.hidden=true;publishBox.hidden=true;
    const login=document.createElement('form');login.style.cssText=panel.style.cssText;
    const title=document.createElement('h3');title.textContent='管理者ログイン';login.append(title);
    const makeInput=(name,type)=>{const label=document.createElement('label');label.textContent=name;label.style.cssText='display:block;margin:8px 0';const el=document.createElement('input');el.type=type;el.required=true;el.autocomplete=type==='password'?'current-password':'username';el.style.cssText='display:block;width:100%;padding:8px;border:1px solid #b9c5d3;border-radius:4px';label.append(el);login.append(label);return el;};
    const email=makeInput('メールアドレス','email'),password=makeInput('管理者用パスワード','password');
    const loginButton=document.createElement('button');loginButton.type='submit';loginButton.className='search-button';loginButton.textContent='ログイン';login.append(loginButton);
    const loginStatus=document.createElement('p');loginStatus.setAttribute('role','status');loginStatus.style.cssText='margin-top:8px;overflow-wrap:anywhere';login.append(loginStatus);
    panel.before(login);form.before(publishBox);publishBox.append(publishButton,linkBox);
    const logout=document.createElement('button');logout.type='button';logout.className='reset-button';logout.textContent='ログアウト';publishBox.append(logout);logoutControl=logout;
    logout.addEventListener('click',async()=>{if(publishing||busy)return;if(leagueDraftPending()&&!confirm('6項目の未保存入力だけが失われます。一時保存済みデータは残ります。ログアウトしますか？'))return;try{await request('/auth/v1/logout?scope=local',{method:'POST',token:await adminToken()});}catch{}session=null;dirty=false;location.reload();});
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
            panel.hidden=false;if(leaguePanel)leaguePanel.hidden=false;publishBox.hidden=false;login.hidden=true;
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