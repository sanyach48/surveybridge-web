const fs=require('fs'),vm=require('vm'),assert=require('assert');
const ctx={console,TextEncoder,TextDecoder,Blob,Uint8Array,ArrayBuffer,Date,atob,btoa,setTimeout,clearTimeout,setImmediate};ctx.window=ctx;ctx.self=ctx;vm.createContext(ctx);
for(const file of ['vendor/xlsx.full.min.js','vendor/export-libs.js','engine.js'])vm.runInContext(fs.readFileSync('dist/'+file,'utf8'),ctx);
const E=vm.runInContext('BrowserEngine',ctx),X=ctx.XLSX;
function load(text){return E.load({name:'test.md',data:Buffer.from(text).toString('base64')}).configs}
async function run(){
 const cfg=load('|ID|Q1|Q2|Q3|Text|Date|Clock|Duration|\n|---|---|---|---|---|---|---|---|\n|a|1|1|1|  001  |29.02.2024|13:45|27:30:00|\n|b|2|1|2|002|01.03.2024|00:00|00:01:00|\n|c|999|2|2|hello|02.03.2024|23:59|00:00:00|\n|d||2|2||03.03.2024|12:00|01:00:00|');
 for(const i of [1,2,3])Object.assign(cfg[i],{mode:'number',missing_codes:[999]});
 Object.assign(cfg[4],{mode:'text',label:' Text ',value_labels:{'001':'WRONG'}});
 Object.assign(cfg[5],{mode:'date',input_format:'%d.%m.%Y',date_display:'EDATE10'});
 Object.assign(cfg[6],{mode:'time',input_format:'%H:%M',date_display:'TIME8'});
 Object.assign(cfg[7],{mode:'duration',date_display:'TIME8'});
 const p={configs:cfg,options:{trim:true},selected:['q2'],fields:['percent','valid','count']};
 assert.equal(E.transform(p).cols[4].values[0],'001');assert.equal(E.preview(p).rows[0][4],'001');
 assert.equal(E.transform({...p,options:{trim:false}}).cols[4].values[0],'  001  ');
 assert.equal(E.preview(p).rows[0][5],'29.02.2024');assert.equal(E.preview({...p,codes:true}).rows[0][6],49500);
 assert.equal(E.preview({...p,col:7,col_count:10}).columns.length,1);
 assert.equal(E.checks({...p,kind:'range',min:1,max:1}).count,1);
 assert.equal(E.checks({...p,kind:'uniform',selected:['q2','q3','q4'],threshold:90,minimum:2}).count,3);
 assert.equal(E.checks({...p,kind:'logic',rules:[{left:'q2',left_op:'missing',right:'q3',right_op:'eq',right_value:2}]}).count,2);
 const report=E.distributions(p),r=report.tables[0].rows.find(r=>r[0]==='1');assert.deepEqual(Array.from(r.slice(1)),[25,50,1]);
 assert.equal(E.distributions({...p,selected:['q2','q3'],two_way:true,basis:'count'}).tables[0].rows.at(-1).at(-1),2);
 const zipBlob=await E.exportData(p);const zip=await ctx.ExportLib.JSZip.loadAsync(await zipBlob.arrayBuffer());
 const book=X.read(await zip.file('survey.xlsx').async('uint8array'),{type:'array'});
 assert.equal(book.Sheets.Data.E2.v,'001');assert.equal(book.Sheets.Data.E2.t,'s');assert.equal(book.Sheets.SPSS_Import.E2.v,'~001');
 const syntax=await zip.file('import.sps').async('string');assert(syntax.includes('COMPUTE q5=SUBSTR(q5,2).'));assert(syntax.includes('MISSING VALUES q2 (999).'));assert(syntax.includes('EDATE10'));
 
 for(const fmt of ['xlsx','docx']){const blob=await E.reportFile(report,fmt);assert(blob.size>500);}
 assert.throws(()=>E.parseDate('31.02.2024',cfg[5]));
 const hostile=JSON.parse(JSON.stringify(p));hostile.configs[1].measure='scale). HOST COMMAND=[malicious]';assert.throws(()=>E.transform(hostile));
 const hostileDate=JSON.parse(JSON.stringify(p));hostileDate.configs[5].date_display='EDATE10). HOST COMMAND=[malicious]';assert.throws(()=>E.transform(hostileDate));
 const header=Array.from({length:151},(_,i)=>'Col'+i),synthetic=[header,header.map(()=>'---')];for(let i=0;i<666;i++)synthetic.push(header.map((_,j)=>j===0?String(i+1):String(i%5+1)));const markdown=synthetic.map(r=>'|'+r.join('|')+'|').join('\n');
 const full=E.load({name:'synthetic.md',data:Buffer.from(markdown).toString('base64')});assert.equal(full.rows,666);assert.equal(full.columns,151);const fullZip=await E.exportData({configs:full.configs});assert(fullZip.size>10000);

 console.log('PASS: full 666x151, dates/times, missing codes, preview, range/logic/uniformity, frequencies/crosstabs, XLSX/DOCX/ZIP, String 001 and SPSS syntax.');
}
run().catch(e=>{console.error(e);process.exit(1)});
