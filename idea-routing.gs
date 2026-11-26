
/** Hobby Central — single-spreadsheet categorisation and sync.
 * Telegram polling remains separate in Code.gs.
 * Install triggers only after the old project has been retired.
 */
const HV_CATEGORIES = ['Writing','Art','Professional Branding','Personal Branding'];
const HV_MAX_PER_RUN = 5;
const HV_MAX_RUNTIME_MS = 4 * 60 * 1000;

function hvMasterCols_(sheet) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim().toLowerCase());
  const locate = (...names) => {
    const index = headers.findIndex(h=>names.includes(h));
    if(index<0) throw new Error('Missing All Tasks header: '+names.join('/'));
    return index+1;
  };
  return {
    id:locate('source id'),date:locate('date added'),title:locate('title'),
    summary:locate('summary'),categories:locate('categories','category'),
    status:locate('status'),themes:locate('themes'),
    message:locate('original message'),url:locate('source url'),notes:locate('notes')
  };
}
function hvMasterRecord_(sheet,row) {
  const cols=hvMasterCols_(sheet), values=sheet.getRange(row,1,1,sheet.getLastColumn()).getValues()[0];
  const record={};Object.keys(cols).forEach(k=>record[k]=values[cols[k]-1]);return record;
}
function hvCategoryList_(value) {
  return [...new Set(String(value||'').split(',').map(s=>s.trim()).filter(Boolean))];
}
function hvValidateCategories_(value) {
  const categories=hvCategoryList_(value);
  const invalid=categories.filter(c=>!HV_CATEGORIES.includes(c));
  if(invalid.length)throw new Error('Invalid category: '+invalid.join(', ')+'. Allowed: '+HV_CATEGORIES.join(', '));
  return categories;
}
function hvFindRow_(sheet,id) {
  const n=sheet.getLastRow();if(n<2)return null;
  const ids=sheet.getRange(2,1,n-1,1).getDisplayValues();
  const idx=ids.findIndex(r=>String(r[0])===String(id));return idx<0?null:idx+2;
}
function categoriseIdea(rawMessage) {
  if(!String(rawMessage||'').trim())throw new Error('Message required; run processIdeaVault instead.');
  const key=PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if(!key)throw new Error('OPENAI_API_KEY missing.');
  const instructions=`Classify a personal creative idea. Return ONLY JSON with keys title, summary, categories (array), type, destination, format, medium, themes, notes.
Allowed categories: Writing, Art, Professional Branding, Personal Branding. Multiple categories ONLY for genuinely separate creative work. If unclear return [].
Writing: essays, articles, pitches, fiction, poetry, Substack; NOT ordinary social posts.
Art: tutorials to try, original visual art, references, inspiration. An Instagram reference link alone does not imply Personal Branding.
Professional Branding: LinkedIn and other professional content. Personal Branding: personal Instagram/social content.
For Writing type: Article, Essay, Short Story, Novel, Poetry, Other; destination: Pitch, Self Publish, Undecided.
For Art type: Tutorial, Original, Reference, Inspiration.
Professional format: Text Post, Carousel, Article, Video, Other. Personal format: Reel, Carousel, Photo, Story, Other.
Themes is comma-separated. Medium and notes may be blank. Do not invent publishing intent.`;
  const response=UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions',{
    method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+key},
    payload:JSON.stringify({model:'gpt-4.1-mini',messages:[{role:'system',content:instructions},{role:'user',content:String(rawMessage)}],response_format:{type:'json_object'},temperature:0.2}),muteHttpExceptions:true
  });
  if(response.getResponseCode()!==200)throw new Error('OpenAI API error: '+response.getContentText());
  const idea=JSON.parse(JSON.parse(response.getContentText()).choices[0].message.content);
  idea.categories=Array.isArray(idea.categories)?[...new Set(idea.categories.filter(c=>HV_CATEGORIES.includes(c)))]:[];
  return idea;
}
function processIdeaVault() {
  const lock=LockService.getScriptLock();if(!lock.tryLock(10000))return;
  try {
    const vault=SpreadsheetApp.getActiveSpreadsheet();
    const inbox=vault.getSheetByName('Raw Messages'), master=vault.getSheetByName('All Tasks');
    if(!inbox||!master)throw new Error('Raw Messages or All Tasks missing.');
    const rows=inbox.getDataRange().getValues(), start=Date.now();let count=0;
    for(let i=1;i<rows.length;i++) {
      const [sourceId,date,chat,message,source,url,status]=rows[i];
      if(!sourceId||!message||['Routed','Needs Review'].includes(String(status||'').trim()))continue;
      if(count>=HV_MAX_PER_RUN||Date.now()-start>=HV_MAX_RUNTIME_MS)break;count++;
      try {
        let row=hvFindRow_(master,sourceId),idea;
        if(row) {
          const saved=hvMasterRecord_(master,row);
          // Existing records: don't reclassify, and don't use obsolete Destination.
          const categories=hvCategoryList_(saved.categories).filter(c=>HV_CATEGORIES.includes(c));
          if(!categories.length) {
            if(String(saved.status)==='Needs Review')inbox.getRange(i+1,7).setValue('Needs Review');
            else throw new Error('Existing idea has legacy/invalid categories; review All Tasks row '+row);
            continue;
          }
          idea={title:saved.title,summary:saved.summary,categories,themes:saved.themes,notes:saved.notes,type:'',destination:'',format:'',medium:''};
        } else {
          idea=categoriseIdea(String(message));
          const cols=hvMasterCols_(master),data=new Array(master.getLastColumn()).fill('');
          const record={id:String(sourceId),date,title:idea.title||'',summary:idea.summary||'',categories:idea.categories.join(', '),status:idea.categories.length?'Idea':'Needs Review',themes:idea.themes||'',message,url:url||'',notes:idea.notes||''};
          Object.keys(cols).forEach(k=>data[cols[k]-1]=record[k]);
          master.getRange(master.getLastRow()+1,1,1,data.length).setValues([data]);
        }
        if(!idea.categories.length){inbox.getRange(i+1,7).setValue('Needs Review');continue;}
        for(const category of idea.categories)hvEnsureDestination_(vault,category,String(sourceId),date,message,url,idea);
        inbox.getRange(i+1,7).setValue('Routed');Logger.log(sourceId+' -> '+idea.categories.join(', '));
      } catch(error){inbox.getRange(i+1,7).setValue('Failed');Logger.log('Failed inbox row '+(i+1)+': '+error.stack);}
    }
  } finally{lock.releaseLock();}
}
function hvEnsureDestination_(vault,category,id,date,message,url,idea) {
  const sheet=vault.getSheetByName(category);if(!sheet)throw new Error('Missing tab: '+category);
  if(hvFindRow_(sheet,id))return;
  const cats=(idea.categories||[]).join(', ');
  if(category==='Writing'){
    const type=['Article','Essay','Short Story','Novel','Poetry','Other'].includes(idea.type)?idea.type:'Other';
    const dest=['Pitch','Self Publish','Undecided'].includes(idea.destination)?idea.destination:'Undecided';
    sheet.appendRow([id,date,idea.title||'',type,cats,dest,'Idea',idea.summary||'',idea.themes||'',message||'',url||'',idea.notes||'','','','','','','']);
  } else if(category==='Art'){
    const type=['Tutorial','Original','Reference','Inspiration'].includes(idea.type)?idea.type:'Original';
    sheet.appendRow([id,date,idea.title||'',type,cats,idea.medium||'',['Reference','Inspiration'].includes(type)?'Saved':'To Try',idea.summary||'',idea.themes||'',message||'',url||'',idea.notes||'']);
  } else {
    const extra=['Format: '+(idea.format||'Other'),'Categories: '+cats,idea.notes||''].filter(Boolean).join('\n');
    sheet.appendRow([id,date,idea.title||'',idea.summary||'','Idea',idea.themes||'',message||'',url||'',extra]);
  }
}
function setupHobbyVaultSync(){
  const vault=SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='syncHobbyVaultEdit').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncHobbyVaultEdit').forSpreadsheet(vault).onEdit().create();
  Logger.log('Edit trigger installed.');
}
function hvFieldMaps_(master){
  const m=hvMasterCols_(master);
  return {'All Tasks':{title:m.title,summary:m.summary,themes:m.themes,message:m.message,url:m.url,notes:m.notes},
    'Writing':{title:3,summary:8,themes:9,message:10,url:11,notes:12},
    'Art':{title:3,summary:8,themes:9,message:10,url:11,notes:12},
    'Professional Branding':{title:3,summary:4,themes:6,message:7,url:8,notes:9},
    'Personal Branding':{title:3,summary:4,themes:6,message:7,url:8,notes:9}};
}
function syncHobbyVaultEdit(e){
  if(!e||!e.range)return;
  const lock=LockService.getScriptLock();if(!lock.tryLock(5000))return;
  try{
    const sheet=e.range.getSheet(),name=sheet.getName();
    if(!['All Tasks',...HV_CATEGORIES].includes(name)||e.range.getLastRow()<2)return;
    const vault=sheet.getParent(),master=vault.getSheetByName('All Tasks'),maps=hvFieldMaps_(master),sourceMap=maps[name];
    const catCol=hvMasterCols_(master).categories;
    for(let row=Math.max(2,e.range.getRow());row<=e.range.getLastRow();row++){
      const sourceId=String(sheet.getRange(row,1).getDisplayValue());if(!sourceId)continue;
      if(name==='All Tasks'&&e.range.getColumn()<=catCol&&e.range.getLastColumn()>=catCol)hvSyncCategoryMembership_(vault,row);
      for(const [field,col] of Object.entries(sourceMap)){
        if(col<e.range.getColumn()||col>e.range.getLastColumn())continue;
        if(field==='notes'&&['Personal Branding','Professional Branding'].includes(name))continue;
        const value=sheet.getRange(row,col).getValue();
        for(const [other,otherMap] of Object.entries(maps)){
          if(other===name||!otherMap[field])continue;
          if(field==='notes'&&['Personal Branding','Professional Branding'].includes(other))continue;
          const dest=vault.getSheetByName(other),targetRow=dest&&hvFindRow_(dest,sourceId);
          if(targetRow)dest.getRange(targetRow,otherMap[field]).setValue(value);
        }
      }
      hvRefreshLive_(vault,sourceId);
    }
  }finally{lock.releaseLock();}
}
function hvSyncCategoryMembership_(vault,masterRow){
  const master=vault.getSheetByName('All Tasks'),r=hvMasterRecord_(master,masterRow);
  if(!r.id)return;
  const categories=hvValidateCategories_(r.categories);
  // Validate all tabs before making any changes.
  for(const c of HV_CATEGORIES)if(!vault.getSheetByName(c))throw new Error('Missing tab: '+c);
  let archive=null;
  for(const category of HV_CATEGORIES){
    const sheet=vault.getSheetByName(category),active=hvFindRow_(sheet,r.id),wanted=categories.includes(category);
    if(wanted&&!active){
      archive=archive||hvGetCategoryArchive_(vault);
      if(!hvRestoreCategory_(archive,sheet,r.id,category))hvEnsureDestination_(vault,category,String(r.id),r.date,r.message,r.url,{title:r.title,summary:r.summary,categories,type:'',destination:'',format:'',medium:'',themes:r.themes,notes:r.notes});
    } else if(!wanted&&active){
      archive=archive||hvGetCategoryArchive_(vault);
      hvArchiveCategory_(archive,sheet,active,r.id,category);
    }
  }
  for(const category of categories){
    const sheet=vault.getSheetByName(category),row=hvFindRow_(sheet,r.id);
    if(row&&['Writing','Art'].includes(category))sheet.getRange(row,5).setValue(categories.join(', '));
  }
  Logger.log('Membership synced: '+r.id+' -> '+categories.join(', '));
}
function hvGetCategoryArchive_(vault){
  let sheet=vault.getSheetByName('_Category Archive');
  if(!sheet){sheet=vault.insertSheet('_Category Archive');sheet.getRange(1,1,1,5).setValues([['Source ID','Category','Archived At','Original Column Count','Saved Row JSON']]);sheet.setFrozenRows(1);sheet.hideSheet();}
  return sheet;
}
function hvArchiveCategory_(archive,sheet,row,id,category){
  const values=sheet.getRange(row,1,1,sheet.getLastColumn()).getValues()[0];
  const saved=values.map(v=>v instanceof Date?{__hobbyVaultDate:v.toISOString()}:v);
  archive.appendRow([String(id),category,new Date(),saved.length,JSON.stringify(saved)]);
  SpreadsheetApp.flush();sheet.deleteRow(row);
}
function hvRestoreCategory_(archive,target,id,category){
  if(hvFindRow_(target,id))return true;
  const n=archive.getLastRow();if(n<2)return false;
  const rows=archive.getRange(2,1,n-1,5).getValues();
  for(let i=rows.length-1;i>=0;i--){
    if(String(rows[i][0])!==String(id)||rows[i][1]!==category)continue;
    const saved=JSON.parse(rows[i][4]).map(v=>v&&typeof v==='object'&&v.__hobbyVaultDate?new Date(v.__hobbyVaultDate):v);
    target.getRange(target.getLastRow()+1,1,1,saved.length).setValues([saved]);return true;
  }
  return false;
}
function refreshHobbyVaultLive(){
  const vault=SpreadsheetApp.getActiveSpreadsheet();
  const ids=new Set();for(const name of HV_CATEGORIES){const s=vault.getSheetByName(name);if(!s||s.getLastRow()<2)continue;s.getRange(2,1,s.getLastRow()-1,1).getDisplayValues().forEach(r=>{if(r[0])ids.add(r[0]);});}
  ids.forEach(id=>hvRefreshLive_(vault,id));
}
function hvRefreshLive_(vault,id){
  const live=vault.getSheetByName('Live');if(!live)return;
  for(const name of HV_CATEGORIES){
    const s=vault.getSheetByName(name),row=s&&hvFindRow_(s,id);if(!row)continue;
    const v=s.getRange(row,1,1,name==='Writing'?18:name==='Art'?12:9).getValues()[0];
    const status=String(v[name==='Writing'||name==='Art'?6:4]||'').trim();
    const done=status==='Published'||(name==='Art'&&status==='Completed'&&!['Reference','Inspiration'].includes(v[3]));
    if(!done)continue;
    const existing=live.getLastRow()<2?[]:live.getRange(2,1,live.getLastRow()-1,3).getDisplayValues();
    if(existing.some(r=>r[0]===String(id)&&r[2]===name))continue;
    live.appendRow([id,v[2],name,name==='Writing'||name==='Art'?v[3]:'Social Content','',name==='Writing'?(v[12]||v[5]):name==='Professional Branding'?'LinkedIn':name==='Personal Branding'?'Instagram':'','','']);
  }
}
function previewIdeaVaultInbox(){
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Messages'),rows=s.getDataRange().getValues();
  Logger.log('Messages: '+(rows.length-1));rows.slice(1).forEach((r,i)=>Logger.log('Row '+(i+2)+' | '+r[0]+' | '+String(r[3]).slice(0,120)));
}
function testIdeaCategorisation(){
  ['I want to paint a watercolour lotus.','A story about alternate versions of ourselves.','A LinkedIn post about brand naming trends.','Make an Instagram reel of my original watercolour painting.'].forEach(m=>Logger.log(JSON.stringify(categoriseIdea(m))));
}
/** Read-only: checks legacy values, duplicate IDs, and required headers. */
function auditHobbyVaultBeforeDeletingF(){
  const vault=SpreadsheetApp.getActiveSpreadsheet();
  const master=vault.getSheetByName('All Tasks'),cols=hvMasterCols_(master);
  Logger.log('All Tasks header map: '+JSON.stringify(cols));
  const rows=master.getDataRange().getValues(), seen=new Set();let legacy=0,duplicates=0;
  for(let i=1;i<rows.length;i++){
    const id=String(rows[i][cols.id-1]||'');if(!id)continue;
    if(seen.has(id)){duplicates++;Logger.log('Duplicate ID row '+(i+1)+': '+id);}seen.add(id);
    const raw=String(rows[i][cols.categories-1]||'');
    const invalid=hvCategoryList_(raw).filter(c=>!HV_CATEGORIES.includes(c));
    if(invalid.length){legacy++;Logger.log('Legacy category row '+(i+1)+': '+raw);}
  }
  Logger.log('Audit: '+legacy+' legacy/invalid category rows; '+duplicates+' duplicate IDs. No data changed.');
}
function testHobbyVaultCategories(){auditHobbyVaultBeforeDeletingF();}