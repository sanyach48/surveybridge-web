/* SurveyBridge browser engine. Does not send files or answers over the network. */
const BrowserEngine=(()=>{
 let raw=[],headers=[];
 const dateModes=['date','datetime','time','duration'];
 const epoch=Date.UTC(1582,9,14);
 const text=x=>x===null||x===undefined||String(x).trim()===''?'':String(x);
 const num=x=>x===null||x===undefined||String(x).trim()===''?null:(Number.isFinite(Number(String(x).replace(',','.')))?Number(String(x).replace(',','.')):null);
 const norm=(x,trim)=>trim?text(x).trim():text(x);
 function parseDate(s,c){
  if(c.mode==='duration'){
   const m=String(s).trim().match(/^(-?)(\d+):(\d{2})(?::(\d{2}))?$/);
   if(!m||+m[3]>59||+(m[4]||0)>59)throw Error('ожидается ЧЧ:ММ:СС');
   return (m[1]?-1:1)*(+m[2]*3600 + +m[3]*60 + +(m[4]||0));
  }
  const specs={'%Y':['(\\d{4})','year'],'%y':['(\\d{2})','shortyear'],'%m':['(\\d{1,2})','month'],'%d':['(\\d{1,2})','day'],'%H':['(\\d{1,2})','hour'],'%I':['(\\d{1,2})','hour12'],'%M':['(\\d{1,2})','minute'],'%S':['(\\d{1,2})','second'],'%p':['(AM|PM)','ampm']};
  let keys=[],pattern='',f=c.input_format||'%Y-%m-%d';
  for(let i=0;i<f.length;i++){const code=f.slice(i,i+2);if(specs[code]){pattern+=specs[code][0];keys.push(specs[code][1]);i++}else pattern+=f[i].replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}
  const m=String(s).trim().match(new RegExp('^'+pattern+'$','i'));if(!m)throw Error('не соответствует формату');
  const v={year:1900,month:1,day:1,hour:0,minute:0,second:0};keys.forEach((k,i)=>v[k]=k==='ampm'?m[i+1].toUpperCase():+m[i+1]);
  if(v.shortyear!==undefined)v.year=v.shortyear+(v.shortyear<=68?2000:1900);
  if(v.hour12!==undefined){if(v.hour12<1||v.hour12>12)throw Error('часы');v.hour=v.hour12%12+(v.ampm==='PM'?12:0)}
  if(v.month<1||v.month>12||v.day<1||v.day>31||v.hour<0||v.hour>23||v.minute>59||v.second>59)throw Error('недопустимая дата');
  const d=new Date(0);d.setUTCFullYear(v.year,v.month-1,v.day);d.setUTCHours(v.hour,v.minute,v.second,0);
  if(d.getUTCFullYear()!==v.year||d.getUTCMonth()!==v.month-1||d.getUTCDate()!==v.day)throw Error('недопустимая дата');
  return c.mode==='time'?v.hour*3600+v.minute*60+v.second:(d.getTime()-epoch)/1000;
 }
 function load(p){
  if(p.reference)throw Error('Образец SAV не поддерживается в браузерной версии. Используйте JSON-настройки.');
  const bytes=Uint8Array.from(atob(p.data),c=>c.charCodeAt(0));let rows,sheets=[];
  if(bytes.byteLength>20*1024*1024)throw Error('Размер файла превышает 20 МБ.');
  if(p.name.toLowerCase().endsWith('.md')){
   rows=new TextDecoder('utf-8').decode(bytes).replace(/^\uFEFF/,'').split(/\r?\n/).map(l=>l.trim()).filter(l=>l.startsWith('|')).map(l=>l.slice(1,-1).split(/(?<!\\)\|/).map(v=>v.replace(/\\\|/g,'|')));
   if(rows.length<3||!rows[1].every(v=>/^:?-+:?$/.test(v.trim())))throw Error('Не найдена Markdown-таблица.');rows.splice(1,1);
  }else{
   const book=XLSX.read(bytes,{type:'array',cellDates:true,raw:true,UTC:true});sheets=book.SheetNames;
   const sh=book.Sheets[sheets[p.sheet||0]];if(!sh)throw Error('Нет листа с таким номером.');
   const range=XLSX.utils.decode_range(sh['!ref']||'A1');
   if(range.e.c>=2000||range.e.r>=50000||(range.e.c+1)*(range.e.r+1)>2000000)throw Error('Слишком большая таблица: максимум 50 000 строк, 2000 столбцов и 2 млн ячеек.');
   rows=XLSX.utils.sheet_to_json(sh,{header:1,raw:true,defval:'',blankrows:false});rows=rows.slice(p.header||0);
   rows=rows.map(r=>r.map(v=>v instanceof Date?v.toISOString().slice(0,19).replace('T',' '):text(v)));
  }
  if(rows.length<2)throw Error('В таблице нет респондентов.');
  const width=rows[0].length;if(rows.some(r=>r.length>width))throw Error('Есть данные за пределами заголовков. Проверьте строку заголовков.');
  if(width>2000||rows.length>50000||width*rows.length>2000000)throw Error('Таблица превышает допустимый объём.');
  headers=rows[0].map((v,i)=>text(v)||'Столбец '+(i+1));raw=rows.slice(1).map(r=>Array.from({length:width},(_,i)=>text(r[i])));
  const configs=headers.map((h,j)=>{
   const frequencies=Object.create(null);raw.forEach(r=>{const v=text(r[j]);frequencies[v]=(frequencies[v]||0)+1});const unique=Object.keys(frequencies).filter(Boolean).sort();
   const numeric=unique.length>0&&unique.every(v=>num(v)!==null);let mode=numeric?'number':unique.length>0&&unique.length<=30?'category':'text';
   const mapping=Object.create(null);if(mode==='category')unique.forEach((v,i)=>mapping[v]=i+1);
   let group='',measure=numeric?'scale':'nominal';
   if(h.includes(' / ')&&unique.length===1&&!/другое|other/i.test(h)){mode='checkbox';group=h.split(' / ')[0]}
   else if(unique.length&&unique.every(v=>/^\d+\s+место$/.test(v))){mode='category';measure='ordinal';unique.forEach(v=>mapping[v]=parseInt(v))}
   return {index:j,source:h,name:'q'+(j+1),label:h,include:true,mode,measure,mapping,value_labels:{},blank:'missing',missing_tokens:[],missing_codes:[],reference_missing:[],group,present:1,absent:0,conflicts:[],reference:'',discrepancies:0,frequencies,input_format:'%Y-%m-%d',date_display:'EDATE10',decimals:mode==='number'?2:0};
  });
  return {rows:raw.length,columns:width,configs,notes:['Файл прочитан в браузере. После закрытия вкладки загрузите его снова.'],sheets,reference:false,positional:false};
 }
 function transform(p){
  const trim=p.options?.trim!==false,configs=p.configs.filter(c=>c.include),seen=new Set(),errors=[],cols=[];
  if(!raw.length)throw Error('Загрузите данные.');if(!configs.length)throw Error('Не выбраны переменные.');
  for(const c of configs){
   if(!['number','category','text','checkbox',...dateModes].includes(c.mode))throw Error('Недопустимый способ обработки.');
   if(!['nominal','ordinal','scale'].includes(c.measure))throw Error('Недопустимый уровень измерения.');
   const formats={date:['EDATE10','SDATE10','ADATE10','DATE11'],datetime:['DATETIME20','YMDHMS19'],time:['TIME8','TIME5'],duration:['TIME8','DTIME11']};
   if(dateModes.includes(c.mode)&&!formats[c.mode].includes(c.date_display))throw Error('Недопустимый формат даты/времени.');
   if(c.decimals!==undefined&&(!Number.isInteger(+c.decimals)||+c.decimals<0||+c.decimals>15))throw Error('Число десятичных знаков: от 0 до 15.');
   if(!/^[A-Za-z][A-Za-z0-9_]{0,63}$/.test(c.name)||seen.has(c.name.toLowerCase()))errors.push(c.name+': имя должно быть уникальным и содержать латинские буквы, цифры или _.');seen.add(c.name.toLowerCase());
   if(!Number.isInteger(c.index)||c.index<0||c.index>=headers.length)throw Error('Некорректный столбец в настройках.');
   const mapping=new Map();for(const [v,code] of Object.entries(c.mapping||{})){const k=norm(v,trim);if(c.mode==='category'&&mapping.has(k)&&num(mapping.get(k))!==num(code))errors.push(c.name+': после удаления пробелов ответ имеет разные коды.');mapping.set(k,code)}
   const codes=c.missing_codes||[];if(codes.length>3||codes.some(x=>num(x)===null))errors.push(c.name+': укажите до трёх числовых кодов пропуска.');
   const tokens=new Set((c.missing_tokens||[]).map(v=>norm(v,trim)));
   const values=raw.map((r,i)=>{
    const x=norm(r[c.index],trim);if(tokens.has(x))return c.mode==='text'?'':null;
    if(x===''){
     if(c.mode==='text')return '';
     if(c.blank==='reference'){errors.push(c.name+': правило «как в образце» недоступно. Выберите пустую ячейку явно.');return null}
     return c.blank==='zero'?0:c.blank==='missing'?null:num(c.blank);
    }
    if(c.mode==='text')return x;
    if(c.mode==='checkbox')return 1;
    if(dateModes.includes(c.mode)){try{return parseDate(x,c)}catch(e){errors.push(`${c.name}, строка ${i+1}: дата/время не соответствует формату`);return null}}
    const v=c.mode==='number'?num(x):num(mapping.get(x));if(v===null)errors.push(`${c.name}, строка ${i+1}: ${c.mode==='number'?'не число':'нет кода для «'+x.slice(0,40)+'»'}`);return v;
   });
   const labels=Object.create(null);if(c.mode!=='text')for(const [k,v] of Object.entries(c.value_labels||{}))if(num(k)!==null)labels[String(num(k))]=v;
   if(c.mode==='category')for(const [v,code] of mapping){if(num(code)!==null&&labels[String(num(code))]===undefined)labels[String(num(code))]=norm(v,trim)}
   if(c.mode==='checkbox'){labels['0']='Не выбрано';labels['1']='Выбрано'}
   for(const k of Object.keys(labels))labels[k]=norm(labels[k],trim);
   cols.push({c,values,labels,label:norm(c.label,trim),missing:codes.map(num)});
  }
  if(errors.length)throw Error(errors.slice(0,25).join('\n')+(errors.length>25?'\nВсего ошибок: '+errors.length:''));
  return {cols,count:raw.length};
 }
 const missing=(col,v)=>v===null||v===''||col.missing.includes(v);
 function display(col,v,codes=false){
  if(v===null||v==='')return null;if(codes)return v;
  if(col.labels[String(v)]!==undefined)return col.labels[String(v)];
  const c=col.c;if(dateModes.includes(c.mode)){
   if(c.mode==='time'||c.mode==='duration'){const sec=Math.abs(Math.trunc(v));return (v<0?'-':'')+[Math.floor(sec/3600),Math.floor(sec/60)%60,sec%60].map(x=>String(x).padStart(2,'0')).join(':')}
   const d=new Date(epoch+v*1000),iso=d.toISOString(),y=iso.slice(0,4),m=iso.slice(5,7),day=iso.slice(8,10);
   if(c.mode==='datetime')return iso.slice(0,19).replace('T',' ');
   return c.date_display==='ADATE10'?`${m}/${day}/${y}`:c.date_display==='SDATE10'?`${y}/${m}/${day}`:c.date_display==='DATE11'?`${day}-${['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'][+m-1]}-${y}`:`${day}.${m}.${y}`;
  }return v;
 }
 function preview(p){const d=transform(p),row=Math.max(0,+p.row||0),col=Math.max(0,+p.col||0),rc=Math.max(1,Math.min(100,+p.row_count||10)),cc=Math.max(1,Math.min(150,+p.col_count||10)),cols=d.cols.slice(col,col+cc);return {columns:cols.map(x=>({name:x.c.name,label:x.label})),rows:Array.from({length:Math.max(0,Math.min(rc,d.count-row))},(_,i)=>cols.map(x=>display(x,x.values[row+i],p.codes))),row,col,total_rows:d.count,total_columns:d.cols.length}}
 function selection(p,d){if(!p.selected?.length)throw Error('Выберите переменные.');return p.selected.map(n=>{const c=d.cols.find(c=>c.c.name===n);if(!c)throw Error('Переменная исключена или переименована: '+n);return c})}
 function checks(p){
  const d=transform(p),rows=[];
  const issue=(i,kind,n,v,reason)=>rows.push([i+1,text(raw[i][0]),kind,n,v,reason]);
  if(p.kind==='range'){
   const lo=num(p.min),hi=num(p.max);if(lo===null||hi===null||lo>hi)throw Error('Укажите корректный диапазон.');
   for(const col of selection(p,d)){if(col.c.mode==='text')throw Error('Для диапазона нужны числовые переменные.');col.values.forEach((v,i)=>{if(!missing(col,v)&&(v<lo||v>hi))issue(i,'Диапазон',col.c.name,v,`Вне [${lo}; ${hi}]`)})}
  }else if(p.kind==='uniform'){
   const cols=selection(p,d),minimum=+p.minimum,threshold=+p.threshold;
   if(!Number.isInteger(minimum)||minimum<2||cols.length<minimum||!(threshold>0&&threshold<=100))throw Error('Проверьте порог и минимум валидных ответов.');
   if(cols.some(c=>!['number','category','checkbox'].includes(c.c.mode)))throw Error('Выберите вопросы без текста и дат.');
   for(let i=0;i<d.count;i++){const values=cols.map(c=>c.values[i]).filter((v,j)=>!missing(cols[j],v));if(values.length<minimum)continue;const freq=new Map();values.forEach(v=>freq.set(v,(freq.get(v)||0)+1));const [v,n]=[...freq].sort((a,b)=>b[1]-a[1])[0];if(100*n/values.length>=threshold)issue(i,'Однообразие',cols.map(c=>c.c.name).join(', '),v,`${n} из ${values.length} валидных ответов (${(100*n/values.length).toFixed(1)}%) имеют один код`)}
  }else if(p.kind==='logic'){
   if(!p.rules?.length)throw Error('Добавьте правило.');
   function condition(n,op,x,i){const col=d.cols.find(c=>c.c.name===n);if(!col)throw Error('Нет переменной '+n);const v=col.values[i],m=missing(col,v);if(op==='missing')return m;if(op==='present')return !m;if(m)return false;const target=col.c.mode==='text'?String(x):num(x);if(target===null)throw Error(n+': укажите числовой код.');return op==='eq'?v===target:op==='ne'?v!==target:op==='gt'?v>target:op==='lt'?v<target:false}
   p.rules.forEach((r,ri)=>{for(let i=0;i<d.count;i++)if(condition(r.left,r.left_op,r.left_value,i)&&condition(r.right,r.right_op,r.right_value,i))issue(i,'Логика',r.left+' → '+r.right,d.cols.find(c=>c.c.name===r.right).values[i],'Запрещённое сочетание, правило '+(ri+1))});
  }else throw Error('Неизвестная проверка.');
  return {columns:['Строка','ID из первого столбца','Проверка','Переменная','Значение / код','Причина'],rows,count:rows.length,note:'Результаты требуют проверки исследователем. Строки не удаляются; пропуски исключены из диапазона и однообразия. Однообразие сравнивает одинаковые числовые коды.'};
 }
 const ordered=values=>[...new Set(values)].sort((a,b)=>typeof a==='number'&&typeof b==='number'?a-b:String(a).localeCompare(String(b),'ru'));
 const pct=(n,d)=>d?Math.round(n/d*10000)/100:null;
 function distributions(p){
  const d=transform(p),cols=selection(p,d),tables=[];
  if(p.two_way){
   if(cols.length!==2||cols[0]===cols[1])throw Error('Выберите два разных вопроса.');let [a,b]=cols;if(p.transpose)[a,b]=[b,a];
   const valid=Array.from({length:d.count},(_,i)=>i).filter(i=>!missing(a,a.values[i])&&!missing(b,b.values[i]));const av=ordered(valid.map(i=>a.values[i])),bv=ordered(valid.map(i=>b.values[i]));
   const matrix=av.map(x=>bv.map(y=>valid.filter(i=>a.values[i]===x&&b.values[i]===y).length)),rt=matrix.map(r=>r.reduce((a,b)=>a+b,0)),ct=bv.map((_,j)=>matrix.reduce((s,r)=>s+r[j],0)),basis=p.basis||'count';
   const cell=(n,den)=>basis==='count'?n:pct(n,den);
   const rows=av.map((x,i)=>[display(a,x),...bv.map((_,j)=>cell(matrix[i][j],basis==='row'?rt[i]:basis==='column'?ct[j]:valid.length)),cell(rt[i],basis==='row'?rt[i]:valid.length)]);
   rows.push(['Всего',...ct.map(n=>cell(n,basis==='column'?n:valid.length)),cell(valid.length,valid.length)]);
   tables.push({title:a.label+' × '+b.label,columns:[a.label,...bv.map(v=>String(display(b,v))),'Всего'],rows,note:`Строки: ${a.label}. Столбцы: ${b.label}. Полных пар: ${valid.length} из ${d.count}. Пропуски исключены. Показатель: ${{count:'количество людей',row:'% по строке',column:'% по столбцу',total:'% от всех полных пар'}[basis]}.`});
  }else{
   const fields=p.fields||['percent','valid','count'];if(!fields.length||fields.some(f=>!['percent','valid','count'].includes(f)))throw Error('Выберите столбцы распределения.');
   for(const col of cols){
    const valid=col.values.filter(v=>!missing(col,v)).length;const values=ordered([...col.values.filter(v=>v!==null&&v!==''),...Object.keys(col.labels).map(v=>col.c.mode==='text'?v:num(v)).filter(v=>v!==null)]);
    function row(label,n,miss=false,total=false){const v={percent:pct(n,d.count),valid:miss?null:total?(valid?100:null):pct(n,valid),count:n};return [label,...fields.map(f=>v[f])]}
    const rows=values.map(v=>row(String(display(col,v))+(missing(col,v)?' [пропуск]':''),col.values.filter(x=>x===v).length,missing(col,v)));
    const empty=col.values.filter(v=>v===null||v==='').length;if(empty)rows.push(row('Системный пропуск / пусто',empty,true));rows.push(row('Всего',d.count,false,true));
    tables.push({title:col.label,columns:['Вопрос / вариант ответа',...fields.map(f=>({percent:'% ответов',valid:'% валидных ответов',count:'Количество людей'}[f]))],rows,note:`Всего респондентов: ${d.count}. Валидных: ${valid}. % ответов — от всех; % валидных — без пропусков. Без весов.`});
   }
  }return {tables};
 }
 function workbook(tables){const book=XLSX.utils.book_new();tables.forEach((tab,i)=>{const sheet=XLSX.utils.aoa_to_sheet([[tab.title],[tab.note],tab.columns,...tab.rows]);sheet['!cols']=tab.columns.map((_,j)=>({wch:j===0?65:22}));sheet['!rows']=[{hpt:36},{hpt:55},{hpt:34}];sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:tab.columns.length-1}},{s:{r:1,c:0},e:{r:1,c:tab.columns.length-1}}];XLSX.utils.book_append_sheet(book,sheet,'Распределение '+(i+1))});return XLSX.write(book,{bookType:'xlsx',type:'array',compression:true})}
 async function reportFile(report,format){
  if(format==='xlsx')return new Blob([workbook(report.tables)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const D=ExportLib.Docx,children=[new D.Paragraph({text:'Распределения ответов',heading:D.HeadingLevel.TITLE})];
  for(const t of report.tables){children.push(new D.Paragraph({text:t.title,heading:D.HeadingLevel.HEADING_1}),new D.Paragraph(t.note));for(let start=1;start<t.columns.length;start+=7){const ids=[0,...Array.from({length:Math.min(7,t.columns.length-start)},(_,j)=>start+j)];children.push(new D.Table({width:{size:100,type:D.WidthType.PERCENTAGE},rows:[t.columns,...t.rows].map((r,i)=>new D.TableRow({tableHeader:i===0,children:ids.map(j=>new D.TableCell({children:[new D.Paragraph({children:[new D.TextRun({text:r[j]===null?'':String(r[j]),bold:i===0})]})]}))}))}))}}
  return D.Packer.toBlob(new D.Document({sections:[{properties:{page:{size:{orientation:D.PageOrientation.LANDSCAPE}}},children}]}));
 }
 const quote=s=>"'"+String(s).replace(/[\r\n\t]/g,' ').replace(/'/g,"''")+"'";
 function spss(p,d){
  const lines=["* SurveyBridge: распакуйте комплект в C:\\SurveyBridge или измените CD ниже.","SET UNICODE=ON.","SET DECIMAL=DOT.","CD 'C:\\SurveyBridge'.","GET DATA /TYPE=XLSX /FILE='survey.xlsx' /SHEET=NAME 'SPSS_Import' /CELLRANGE=FULL /READNAMES=ON /DATATYPEMIN PERCENTAGE=100 /LEADINGSPACES IGNORE=NO /TRAILINGSPACES IGNORE=NO."];
  const groups=new Map();
  for(const col of d.cols){const c=col.c,n=c.name;
   if(c.mode==='text'){
    const bytes=Math.max(1,...col.values.map(v=>new TextEncoder().encode(v||'').length));if(bytes>32766)throw Error(n+': текст длиннее допустимой строки импорта SPSS (32766 байт).');
    lines.push(`COMPUTE ${n}=SUBSTR(${n},2).`,`ALTER TYPE ${n} (A${bytes}).`);
   }else{
    lines.push(`ALTER TYPE ${n} (F40.12).`);
    const fmt=dateModes.includes(c.mode)?c.date_display:`F20.${Math.max(0,Math.min(15,+c.decimals||0))}`;
    lines.push(`FORMATS ${n} (${fmt}).`);
    if(col.missing.length)lines.push(`MISSING VALUES ${n} (${col.missing.join(', ')}).`);
    const labels=Object.entries(col.labels);if(labels.length)lines.push(`VALUE LABELS ${n}\n`+labels.map(([k,v])=>` ${num(k)} ${quote(v)}`).join('\n')+'.');
   }
   lines.push(`VARIABLE LABELS ${n} ${quote(col.label)}.`,`VARIABLE LEVEL ${n} (${c.mode==='text'?'NOMINAL':c.measure.toUpperCase()}).`);
   if(c.mode==='checkbox'&&c.group){if(!groups.has(c.group))groups.set(c.group,[]);groups.get(c.group).push(n)}
  }
  let i=0;for(const [label,names] of groups)if(names.length>1)lines.push(`MRSETS /MDGROUP NAME=$set${++i} LABEL=${quote(label)} VARIABLES=${names.join(' ')} VALUE=1.`);
  lines.push('EXECUTE.',"SAVE OUTFILE='survey.sav' /COMPRESSED.");return '\ufeff'+lines.join('\r\n');
 }
 async function exportData(p){
  const d=transform(p),book=XLSX.utils.book_new(),names=d.cols.map(c=>c.c.name);
  const rows=Array.from({length:d.count},(_,i)=>d.cols.map(c=>c.values[i]));
  const sheet=XLSX.utils.aoa_to_sheet([names,...rows]);sheet['!cols']=names.map(()=>({wch:24}));XLSX.utils.book_append_sheet(book,sheet,'Data');
  const importRows=rows.map(r=>r.map((v,j)=>d.cols[j].c.mode==='text'?'~'+(v||''):v));
  XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([names,...importRows]),'SPSS_Import');
  const zip=new ExportLib.JSZip();zip.file('survey.xlsx',XLSX.write(book,{bookType:'xlsx',type:'array',compression:true}));zip.file('import.sps',spss(p,d));
  zip.file('template.json',JSON.stringify({version:4,configs:p.configs,options:p.options,rules:p.rules||[]},null,2));
  zip.file('dictionary.json',JSON.stringify(d.cols.map(c=>({name:c.c.name,label:c.label,mode:c.c.mode,measure:c.c.measure,values:c.labels,missing:c.missing})),null,2));
  zip.file('ПРОЧИТАЙТЕ.txt','1. Распакуйте весь комплект в C:\\SurveyBridge. Или измените строку CD в import.sps на вашу папку.\r\n2. Откройте import.sps в SPSS с поддержкой Unicode и выполните все команды.\r\n3. Будет создан survey.sav.\r\n\r\nЛист Data в Excel содержит подготовленные значения. Служебный лист SPSS_Import содержит маркер ~ перед текстом; команды удаляют маркер. Это сохраняет String, ведущие нули и пустые текстовые ячейки при импорте. Не редактируйте служебный лист вручную. Даты хранятся числом секунд SPSS; формат задаётся командами.\r\nПрямой SAV в браузере не создаётся. Анкета обработана на вашем устройстве.\r\n');
  return zip.generateAsync({type:'blob',compression:'DEFLATE'});
 }
 async function route(path,p){
  let data;if(path==='/load')data=load(p);else if(path==='/preview')data=preview(p);else if(path==='/checks')data=checks(p);else if(path==='/distributions')data=distributions(p);else if(path==='/distributions-export')return {json:async()=>null,blob:async()=>reportFile(distributions(p),p.format)};else if(path==='/export')return {json:async()=>null,blob:async()=>exportData(p)};else throw Error('Неизвестная операция.');
  return {json:async()=>data,blob:async()=>new Blob([JSON.stringify(data)])};
 }
 return {route,load,transform,preview,checks,distributions,reportFile,exportData,parseDate,spss};
})();
