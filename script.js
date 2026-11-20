// STEP 22: Supabaseクラウド保存・管理者認証。
// STEP 17: 閲覧者の案内欄を非表示・試合区分を追加・全条件の一括解除。
// STEP 16: 管理者・閲覧者のローカル分離と更新終了（IndexedDB）。外部公開・認証は未実装。
// STEP 15: 紅白戦CSV追加・ 連動型フィルター・6項目検索（対戦大学追加）・4区分・YouTube再生・リーグ戦チャプター同期（余裕3秒）
// CSVはローカル処理。動画を選択するとYouTubeに接続します。
'use strict';
const $ = id => document.getElementById(id);
const adminMode = new URLSearchParams(location.search).get('admin') === '1';
let dirty = false, busy = 0, publishing = false, ready = false;
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
  "凡打": [
    "ファウル",
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
    const rows=parseCSV(text), headers=(rows.shift()||[]).map(clean);
    const sheet=category==='sheet';
    const required=sheet?['日付','打者','投手','Pitch Type','Velocity']:['日付','打者名','投手名','打者チーム','投手チーム','球種','球速'];
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
const dataByCategory={league:[],practice:[],practiceB:[],intraSquad:[],sheet:[]};
let results=[], page=1, selectedIndex=-1, selectedAngle=null;
const form=$('searchForm');
// 既存HTMLを変更せず検索フォームだけを再構築する。
const filterFields = [['year','年度'],['date','日付'],['category','試合区分'],['opponent','対戦大学'],['pitcher','投手'],['batter','打者'],['result','打撃結果']];
const filterValue = (record,key) => key==='result' ? record.resultClass : record[key];
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
const summary=document.createElement('summary');summary.textContent='CSV読み込み（管理者用）';summary.style.cssText='font-weight:bold;cursor:pointer;margin-bottom:8px;';panel.append(summary);
const hint=document.createElement('p');hint.textContent='更新する区分だけ選択してください。同じ区分は差し替え、他の区分は保持します。更新終了まで閲覧ページには反映されません。';hint.style.marginBottom='12px';panel.append(hint);
for(const category of Object.keys(labels)) {
    const label=document.createElement('label');label.textContent=labels[category];label.style.cssText='display:block;font-weight:bold;margin-top:10px;';
    const input=document.createElement('input');input.type='file';input.accept='.csv,text/csv';input.style.cssText='display:block;width:100%;font-size:11px;margin:6px 0;';label.append(input);panel.append(label);
    const status=document.createElement('p');status.textContent='未読み込み';status.setAttribute('aria-live','polite');status.style.cssText='color:#64748b;overflow-wrap:anywhere;';panel.append(status);uploadStatuses.set(category,status);
    input.addEventListener('change',async()=>{
        const file=input.files[0];if(!file||!ready||publishing)return;
        status.textContent='読み込み中…';input.disabled=true;busy++;updatePublishUI();
        try {
            const records=normalizeCSV(await file.text(),category);
            const remap=category==='league'?remapLeagueAnchors(dataByCategory.league,records):null;
            dataByCategory[category]=records;
            if(remap){leagueAnchors.clear();for(const [k,v] of remap.entries)leagueAnchors.set(k,v);refreshLeaguePanel();}
            markDirty();
            status.textContent=`${records.length.toLocaleString('ja-JP')}件 読み込み完了`+(category==='sheet'?'（所属未確認）':'')+(remap?`｜リーグ戦基準点：${remap.kept}件引継ぎ、${remap.dropped}件要再登録`:'');
            status.style.color='#17633a';
            populateFilters();search();
        } catch(e) {status.textContent='読み込み失敗：'+e.message+'（前回のデータは保持）';status.style.color='#b42318';}
        finally {input.disabled=false;input.value='';busy--;updatePublishUI();}
    });
}
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

// STEP 8: リーグ戦の同期。登録は今回のブラウザタブ内のみ。
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
function setupLeaguePanel() {
    const box=document.createElement('details');box.style.cssText=panel.style.cssText;
    const summary=document.createElement('summary');summary.textContent='リーグ戦 動画登録（管理者用）';summary.style.cssText='font-weight:bold;cursor:pointer;margin-bottom:8px';box.append(summary);
    const p=t=>{const el=document.createElement('p');el.textContent=t;el.style.cssText='margin:8px 0;line-height:1.5;overflow-wrap:anywhere;';box.append(el);return el;};
    p('リーグ戦CSVを読み込んでから登録。基準は投球動作の約2秒前。再生はさらに3秒前から。更新終了を押すと登録内容を保存します。CSV差し替え時は同一と確認できた基準点のみ引き継ぎます。');
    function field(label,el) {const l=document.createElement('label');l.textContent=label;l.style.cssText='display:block;font-weight:bold;margin:10px 0 4px';el.style.cssText='display:block;width:100%;margin-top:4px;padding:5px;border:1px solid #b9c5d3;border-radius:4px;font-size:12px;';l.append(el);box.append(l);return el;}
    const match=field('試合',document.createElement('select'));
    const angle=field('撮影アングル',document.createElement('select'));angle.replaceChildren(...angleOrder.map(a=>new Option(angleLabel(a),a)));
    const url=field('YouTube URL',document.createElement('input'));url.type='url';url.placeholder='https://youtu.be/...';
    const text=field('イニング別チャプター',document.createElement('textarea'));text.rows=5;text.placeholder='05:00 1回表\n18:25 1回裏\n32:10 2回表';
    const status=document.createElement('p');status.setAttribute('aria-live','polite');status.style.cssText='line-height:1.5;margin:8px 0;overflow-wrap:anywhere;';
    function button(label,action) {const b=document.createElement('button');b.type='button';b.textContent=label;b.className='search-button';b.style.cssText='margin:6px 0;font-size:12px;padding:8px;';b.addEventListener('click',()=>{try{if(!ready||publishing)return;action();markDirty();status.style.color='#17633a';}catch(e){status.textContent=e.message;status.style.color='#b42318';}});box.append(b);return b;}
    const rows=()=>dataByCategory.league.filter(r=>leagueMatchKey(r)===match.value);
    const requireMatch=()=>{if(!rows().length)throw new Error('リーグ戦CSVと試合を選択してください。');};
    const getId=()=>{const parsed=parseYouTube({url:url.value,seconds:'0'});if(!parsed)throw new Error('正しいYouTube URLを入力してください。');return parsed.id;};
    const count=()=>rows().filter(r=>Object.keys(playableAngles(r)).length).length;
    function refresh(message) {search();status.textContent=message+'｜この試合の再生位置登録済み：'+count()+' / '+rows().length+'件';updatePitches();}
    button('チャプターを登録・更新',()=>{
        requireMatch();const id=getId(), parsed=chapterLines(text.value), gameRows=rows();
        const candidates=[...parsed.found].map(([inning,start])=>({r:gameRows.find(r=>r.inning===inning),start,inning}));
        const use=candidates.filter(x=>x.r);
        if(!use.length)throw new Error('慶應の攻撃イニングに一致するチャプターがありません。例：05:00 1回表');
        for(const {r,start} of use)if(clockSeconds(r.timestamp)===null)throw new Error(r.inning+' の先頭行の時刻が不正です。');
        for(const {r,start} of use)putAnchor(r,angle.value,{id,start});
        refresh(use.length+'イニング登録。対象外・認識できない行：'+(candidates.length-use.length+parsed.ignored.length)+'行（慶應の守備回等）。');
    });
    p('途中補正・撮影中断：該当イニングと投球を選び、再開地点の動画内時刻を登録。映像がない最初の投球には「動画なし」を登録し、再開した最初の投球に新しい基準を登録します。');
    const inning=field('イニング',document.createElement('select'));
    const pitch=field('基準にする投球（CSV順）',document.createElement('select'));
    const time=field('動画内の時刻（秒または hh:mm:ss）',document.createElement('input'));time.placeholder='1:23:45';
    function updatePitches(){const previous=pitch.value;const a=rows().filter(r=>r.inning===inning.value);pitch.replaceChildren(new Option('選択してください',''),...a.map((r,i)=>{const item=(leagueAnchors.get(anchorSlot(r,angle.value))||[]).find(x=>x.row===r.sourceRow);return new Option((i+1)+'球目｜'+r.timestamp.slice(-8)+'｜'+r.batter+'｜打席内'+r.pitchInAtBat+'球'+(item?item.missing?'【動画なし】':'【基準登録済】':''),String(r.sourceRow));}));if([...pitch.options].some(o=>o.value===previous))pitch.value=previous;}
    function updateInnings(){const previous=inning.value;inning.replaceChildren(new Option('選択してください',''),...[...new Set(rows().map(r=>r.inning).filter(Boolean))].map(i=>new Option(i,i)));if([...inning.options].some(o=>o.value===previous))inning.value=previous;updatePitches();}
    function updateMatches(){const previous=match.value;const m=new Map(dataByCategory.league.map(r=>[leagueMatchKey(r),r.date+' vs '+r.opponent]));match.replaceChildren(new Option('選択してください',''),...[...m].map(([k,label])=>new Option(label,k)));if(m.has(previous))match.value=previous;updateInnings();status.textContent='試合を選択して登録してください。';}
    match.addEventListener('change',()=>{updateInnings();status.textContent='この試合の再生位置登録済み：'+count()+' / '+rows().length+'件';});
    inning.addEventListener('change',updatePitches);angle.addEventListener('change',updatePitches);
    const chosen=()=>{requireMatch();const r=rows().find(r=>r.inning===inning.value&&String(r.sourceRow)===pitch.value);if(!r)throw new Error('基準にする投球を選択してください。');return r;};
    button('この投球から再同期',()=>{const r=chosen(),id=getId(),start=parseTime(time.value);if(start===null||clockSeconds(r.timestamp)===null)throw new Error('動画内時刻またはCSV時刻が不正です。');putAnchor(r,angle.value,{id,start});refresh('補正点を登録しました。');});
    button('この投球から動画なし',()=>{putAnchor(chosen(),angle.value,{missing:true});refresh('中断点を登録しました。再開投球に補正点を登録してください。');});
    button('選択した投球の基準・中断点を削除',()=>{const r=chosen(),key=anchorSlot(r,angle.value);leagueAnchors.set(key,(leagueAnchors.get(key)||[]).filter(a=>a.row!==r.sourceRow));refresh('指定した登録を削除しました。前の基準が再適用される場合があります。');});
    box.append(status);panel.after(box);updateMatches();return updateMatches;
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
async function readPublished(){const db=await openStore();try{return await new Promise((resolve,reject)=>{
    const tx=db.transaction('snapshots','readonly'),req=tx.objectStore('snapshots').get('published');
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
const note=document.createElement('p');note.style.cssText='line-height:1.6;margin-bottom:8px';note.textContent='更新終了を押すまで編集内容は公開されません。';publishBox.append(note);
const publishStatus=document.createElement('p');publishStatus.setAttribute('aria-live','polite');publishStatus.style.cssText='line-height:1.6;overflow-wrap:anywhere;margin:6px 0';publishBox.append(publishStatus);
const publishButton=document.createElement('button');publishButton.type='button';publishButton.className='search-button';publishButton.textContent='更新終了';publishButton.disabled=true;
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
    publishButton.disabled=!session||!ready||busy>0||publishing||!dirty||totalRows()===0;
    if(ready&&!publishing)publishStatus.textContent=`編集中：${totalRows().toLocaleString('ja-JP')}件｜${dirty?'未確定の変更あり':'変更なし'}${lastPublished?'｜最終確定：'+new Date(lastPublished).toLocaleString('ja-JP'):''}`;
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
    const logout=document.createElement('button');logout.type='button';logout.className='reset-button';logout.textContent='ログアウト';publishBox.append(logout);
    logout.addEventListener('click',async()=>{if(publishing||busy)return;if(dirty&&!confirm('未確定の編集を破棄してログアウトしますか？'))return;try{await request('/auth/v1/logout?scope=local',{method:'POST',token:await adminToken()});}catch{}session=null;dirty=false;location.reload();});
    login.addEventListener('submit',async e=>{
        e.preventDefault();if(loginButton.disabled)return;loginButton.disabled=true;loginStatus.textContent='ログイン中…';
        try{
            const value=await request('/auth/v1/token?grant_type=password',{method:'POST',body:{email:email.value,password:password.value}});password.value='';setSession(value);
            const loaded=await readCloud();cloudVersion=loaded.version;
            let snapshot=loaded.snapshot,migration=false;
            if(!snapshot&&localOnly){snapshot=await readPublished();migration=!!snapshot;}
            applySnapshot(snapshot);ready=true;dirty=migration;
            panel.hidden=false;if(leaguePanel)leaguePanel.hidden=false;publishBox.hidden=false;login.hidden=true;
            updatePublishUI();
            if(migration)publishStatus.textContent=`ローカル確定版 ${totalRows().toLocaleString('ja-JP')}件を復元。公開許可を確認後、更新終了を押すと初回クラウド移行します。`;
        }catch(err){session=null;ready=false;password.value='';loginStatus.textContent='ログイン・読み込み失敗：'+err.message;}
        finally{loginButton.disabled=false;}
    });
    publishButton.addEventListener('click',async()=>{
        if(publishButton.disabled)return;
        if(!confirm('ログイン不要で閲覧できるデータとして公開します。部内の許可を得た選手情報・動画・シートBT所属を確認しましたか？'))return;
        publishing=true;publishButton.disabled=true;logout.disabled=true;publishStatus.textContent='保存中…';
        try{
            const snapshot={version:1,publishedAt:new Date().toISOString(),data:dataByCategory,anchors:[...leagueAnchors]};
            // 公開する前に復旧可能なローカル確定版を維持する。
            await savePublished(snapshot,'migration-backup');
            await publishCloud(snapshot);
            try{await savePublished(snapshot);}catch{}
            lastPublished=snapshot.publishedAt;dirty=false;linkBox.hidden=false;
            publishStatus.textContent=`クラウド更新完了：${totalRows().toLocaleString('ja-JP')}件`;
        }catch(e){publishStatus.textContent='更新失敗または結果未確認：'+e.message+'。公開済みデータを確認して再試行してください。ローカルのコピーは保持されます。';}
        finally{publishing=false;logout.disabled=false;publishButton.disabled=!session||!ready||busy>0||!dirty||!totalRows();}
    });
    window.addEventListener('beforeunload',e=>{if(dirty||busy||publishing){e.preventDefault();e.returnValue='';}});
    if(!configured())loginStatus.textContent='まず script.js の公開用キーを設定してください。';
}else loadViewer();