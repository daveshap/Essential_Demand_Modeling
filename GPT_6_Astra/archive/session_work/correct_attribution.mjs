import fs from 'node:fs/promises';
import path from 'node:path';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';
const here=path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/,'$1'));
const original=path.resolve(here,'../outputs/Essential_Demand_ChatGPT_6_Pro_2026-11-16_083431_EST.xlsx');
const target=path.resolve(here,'../outputs/Essential_vs_Derived_Demand_by_Occupation_GPT_6_Astra_2026-11-16_0834.xlsx');
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(original));
console.log((await w.inspect({kind:'region',sheetId:'Method and sources',range:'A5:B6',maxChars:2000})).ndjson);
w.worksheets.getItem('Summary').getRange('A4').values=[['Prepared by GPT-6 Astra in Codex']];
w.worksheets.getItem('Occupations').getRange('A3').values=[['GPT-6 Astra. Created 2026-11-16 08:34:31 EST (America/New_York)']];
w.worksheets.getItem('Method and sources').getRange('A5').values=[['Model']];
w.worksheets.getItem('Method and sources').getRange('B5').values=[['GPT-6 Astra']];
w.worksheets.getItem('Method and sources').getRange('B6').values=[['Prepared by GPT-6 Astra in Codex, with three same-model subagents scoring disjoint occupation sets and a review pass.']];
w.recalculate();
console.log((await w.inspect({kind:'match',searchTerm:'ChatGPT 6 Pro|GPT-6-based|Pro designation',options:{useRegex:true,maxResults:20},maxChars:2000})).ndjson);
for(const [sheetName,range,file]of [['Summary','A2:D5','corrected_summary.png'],['Occupations','A2:D3','corrected_occupation_title.png'],['Method and sources','A5:B8','corrected_provenance.png']]){
 const img=await w.render({sheetName,range,scale:1.5,format:'png'});
 await fs.writeFile(path.join(here,file),new Uint8Array(await img.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(w)).save(target);
await fs.writeFile(path.join(here,'corrected_output_path.txt'),target);
console.log(target);
