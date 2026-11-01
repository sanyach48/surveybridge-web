const esbuild=require('esbuild');
esbuild.build({stdin:{contents:'import * as Docx from "docx";import JSZip from "jszip";window.ExportLib={Docx,JSZip};',resolveDir:process.cwd()},bundle:true,platform:'browser',format:'iife',target:'es2020',outfile:'dist/vendor/export-libs.js',minify:true,legalComments:'eof'}).catch(e=>{console.error(e);process.exit(1)});
