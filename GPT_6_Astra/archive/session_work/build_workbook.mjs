import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

const here=path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/,'$1'));
const read=async name=>JSON.parse((await fs.readFile(path.join(here,name),'utf8')).replace(/^\uFEFF/,''));
const rows=await read('combined.json');
const sourceRows=await read('bls/supplied_source_rows.json');
const meta=await read('bls/supplied_metadata.json');
const expected=await read('independent_summary.json');
const runfile=path.join(here,'run_metadata.json');
let run;
try {run=JSON.parse(await fs.readFile(runfile,'utf8'));} catch {
  const now=new Date();
  const parts=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:'America/New_York',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit',hourCycle:'h23'}).formatToParts(now).map(v=>[v.type,v.value]));
  const local=`${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute}:${parts.second}`;
  run={created_utc:now.toISOString(),created_eastern:local+' EST (America/New_York)',filename:`Essential_Demand_ChatGPT_6_Pro_${parts.year}-${parts.month}-${parts.day}_${parts.hour}${parts.minute}${parts.second}_EST.xlsx`};
  await fs.writeFile(runfile,JSON.stringify(run,null,2));
}
const outdir=path.resolve(here,'../outputs');
await fs.mkdir(outdir,{recursive:true});
const w=Workbook.create();
const summary=w.worksheets.add('Summary');
const occ=w.worksheets.add('Occupations');
const raw=w.worksheets.add('BLS source');
const method=w.worksheets.add('Method and sources');
const ink='#203047',navy='#243D5A',blue='#EEF4FB',gray='#65758B';
const body=(s,range)=>{s.showGridLines=false;s.getRange(range).format={font:{name:'Arial',size:10,color:ink},rowHeight:22,verticalAlignment:'center'};};
const header=(s,range)=>{s.getRange(range).format={fill:navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,horizontalAlignment:'center',verticalAlignment:'center',rowHeight:36};};
const title=(s,cell,text)=>{s.getRange(cell).values=[[text]];s.getRange(cell).format.font={name:'Arial',size:16,bold:true,color:navy};s.getRange(cell).format.rowHeight=28;};
const nums=(s,range,fmt)=>{s.getRange(range).setNumberFormat(fmt);s.getRange(range).format.horizontalAlignment='right';};
const start=6,end=start+rows.length-1;
body(occ,`A1:P${end}`);
title(occ,'A2','Essential and derived demand by occupation');
occ.getRange('A3').values=[['ChatGPT 6 Pro (requested run label). Created '+run.created_eastern]];
occ.getRange('A3').format.font={name:'Arial',size:10,color:gray};
const headings=['SOC / NEM code','Occupation','Employment, 2025 (jobs)','Essential low','Essential median','Essential high','Derived low','Derived median','Derived high','Explanation','Major occupation group','Essential equivalents, low','Essential equivalents, median','Essential equivalents, high','Source workbook row','BLS occupation description'];
occ.getRange('A5:P5').values=[headings];
header(occ,'A5:P5');
const rawStart=5;
const sourceIndex=new Map(sourceRows.map((r,i)=>[r.code,i+rawStart]));
const data=rows.map(r=>[r.code,r.title,null,r.low/100,r.median/100,r.high/100,null,null,null,r.explanation,r.major_title,null,null,null,r.source_row,r.ooh_url||'No occupation-specific link supplied']);
occ.getRange(`A${start}:P${end}`).values=data;
const jobFormula=rows.map(r=>[`=ROUND('BLS source'!D${sourceIndex.get(r.code)}*1000,0)`]);
occ.getRange(`C${start}:C${end}`).formulas=jobFormula;
occ.getRange('G6:I6').formulas=[['=1-F6','=1-E6','=1-D6']];
occ.getRange(`G6:I${end}`).fillDown();
occ.getRange('L6:N6').formulas=[['=C6*D6','=C6*E6','=C6*F6']];
occ.getRange(`L6:N${end}`).fillDown();
const colWidths={A:14,B:56,C:18,D:14,E:14,F:14,G:14,H:14,I:14,J:110,K:54,L:19,M:19,N:19,O:14,P:90};
for(const [c,v] of Object.entries(colWidths)) occ.getRange(`${c}1:${c}${end}`).format.columnWidth=v;
occ.getRange(`A6:P${end}`).format.rowHeight=44;
occ.getRange(`B6:B${end}`).format.wrapText=true;
occ.getRange(`J6:K${end}`).format.wrapText=true;
occ.getRange(`J6:J${end}`).format.verticalAlignment='center';
occ.getRange(`D6:F${end}`).format.fill='#FFF7E6';
nums(occ,`C6:C${end}`,'#,##0');nums(occ,`D6:I${end}`,'0%');nums(occ,`L6:N${end}`,'#,##0');
occ.getRange(`E6:E${end}`).format.font.bold=true;
occ.getRange(`D6:F${end}`).conditionalFormats.addCustom('OR(D6<0,D6>1,$D6>$E6,$E6>$F6)',{fill:'#FDE9E7',font:{color:'#9A241A'}});
occ.freezePanes.freezeRows(5);occ.freezePanes.freezeColumns(2);occ.tabColor=navy;
const ot=occ.tables.add(`A5:P${end}`,true,'OccupationEstimates');ot.showFilterButton=true;

body(raw,`A1:E${sourceRows.length+rawStart-1}`);
title(raw,'A2','BLS Table 1.2 source rows');
raw.getRange('A3').values=[['Source: user-supplied occupation.xlsx, Table 1.2; verified against the official BLS table. Full references: Method and sources.']];
raw.getRange('A4:E4').values=[['NEM code','Occupation title','BLS row type','Employment, 2025 (thousands)','Original workbook row']];
header(raw,'A4:E4');
raw.getRange(`A5:E${sourceRows.length+4}`).values=sourceRows.map(r=>[r.code,r.title,r.type,r.employment_thousands_2025,r.source_row]);
for(const [c,v] of Object.entries({A:14,B:85,C:17,D:27,E:24}))raw.getRange(`${c}1:${c}${sourceRows.length+4}`).format.columnWidth=v;
nums(raw,`D5:D${sourceRows.length+4}`,'#,##0.0');
raw.getRange(`B5:B${sourceRows.length+4}`).format.wrapText=true;
raw.getRange(`A5:E${sourceRows.length+4}`).format.rowHeight=30;
raw.tables.add(`A4:E${sourceRows.length+4}`,true,'BLSSourceRows');raw.freezePanes.freezeRows(4);raw.freezePanes.freezeColumns(2);

body(summary,'A1:F49');summary.tabColor=navy;
for(const[c,v]of Object.entries({A:64,B:21,C:17,D:17,E:17,F:24}))summary.getRange(`${c}1:${c}49`).format.columnWidth=v;
title(summary,'A2','Essential vs. derived demand');
summary.getRange('A3').values=[['2025 U.S. occupational employment baseline']];
summary.getRange('A4').values=[['ChatGPT 6 Pro (requested run label; execution provenance in Method and sources)']];
summary.getRange('A5').values=[['Created '+run.created_eastern]];
summary.getRange('A3:A5').format.font={name:'Arial',size:10,color:gray};
summary.getRange('A7:D7').values=[['Employment-weighted estimate','Low','Median','High']];header(summary,'A7:D7');
summary.getRange('A8:A11').values=[['Essential demand share'],['Derived demand share'],['Essential job-weighted equivalents'],['Derived job-weighted equivalents']];
summary.getRange('B10:D10').formulas=[[`=SUM(Occupations!L${start}:L${end})`,`=SUM(Occupations!M${start}:M${end})`,`=SUM(Occupations!N${start}:N${end})`]];
summary.getRange('B8:D8').formulas=[['=B10/$B$15','=C10/$B$15','=D10/$B$15']];
summary.getRange('B9:D9').formulas=[['=1-D8','=1-C8','=1-B8']];
summary.getRange('B11:D11').formulas=[['=$B$15-D10','=$B$15-C10','=$B$15-B10']];
nums(summary,'B8:D9','0.0%');nums(summary,'B10:D11','#,##0');
summary.getRange('A8:D8').format.fill=blue;summary.getRange('A8:D8').format.font.bold=true;summary.getRange('A8:D11').format.rowHeight=27;
summary.getRange('A14:A18').values=[['Detailed occupations scored'],['Covered employment, 2025 (jobs)'],['BLS published all-occupation total (jobs)'],['Published total less detailed rows (jobs)'],['Coverage of published total']];
summary.getRange('B14:B18').formulas=[[`=COUNTA(Occupations!A${start}:A${end})`],[`=SUM(Occupations!C${start}:C${end})`],[`='BLS source'!D${sourceIndex.get('00-0000')}*1000`],['=B16-B15'],['=B15/B16']];
nums(summary,'B14:B17','#,##0');nums(summary,'B18','0.000%');
summary.getRange('A20').values=[['Estimates are judgments, not measured BLS shares, FTEs or a forecast of surviving jobs.']];summary.getRange('A20').format.font={name:'Arial',size:10,italic:true,color:gray};
summary.getRange('A22:F22').values=[['Major occupation group','Covered jobs','Essential low','Essential median','Essential high','Median essential equivalents']];header(summary,'A22:F22');
const groups=[...new Map(rows.map(r=>[r.major_code,r.major_title])).entries()];
for(let i=0;i<groups.length;i++){
 const rr=23+i;
 summary.getRange(`A${rr}`).values=[[groups[i][1]]];
 summary.getRange(`B${rr}:F${rr}`).formulas=[[
 `=SUMIF(Occupations!$K$${start}:$K$${end},A${rr},Occupations!$C$${start}:$C$${end})`,
 `=SUMIF(Occupations!$K$${start}:$K$${end},A${rr},Occupations!$L$${start}:$L$${end})/B${rr}`,
 `=F${rr}/B${rr}`,
 `=SUMIF(Occupations!$K$${start}:$K$${end},A${rr},Occupations!$N$${start}:$N$${end})/B${rr}`,
 `=SUMIF(Occupations!$K$${start}:$K$${end},A${rr},Occupations!$M$${start}:$M$${end})`]];
 if(i%2===0)summary.getRange(`A${rr}:F${rr}`).format.fill='#F3F6FA';
}
summary.getRange('A23:A44').format.wrapText=true;summary.getRange('A23:F44').format.rowHeight=32;
nums(summary,'B23:B44','#,##0');nums(summary,'C23:E44','0.0%');nums(summary,'F23:F44','#,##0');
summary.getRange('A47:A48').values=[['Occupations with a 0% median essential estimate'],['Occupations with 0% at all three essential estimates']];
summary.getRange('B47:B48').formulas=[[`=COUNTIF(Occupations!E${start}:E${end},0)`],[`=COUNTIF(Occupations!F${start}:F${end},0)`]];nums(summary,'B47:B48','#,##0');

const notes=[
 ['Run label','ChatGPT 6 Pro — requested comparison label.'],
 ['Execution provenance','Prepared in Codex by a GPT-6-based assistant, with three same-model subagents scoring disjoint occupation sets and a review pass. This session does not independently verify a Pro product/tier designation.'],
 ['Created, Eastern time',run.created_eastern],
 ['Created, UTC',run.created_utc],
 ['User definition','Essential demand is “when the human is the entire point” OR nonfungible OR people are willing to pay a premium for the human OR simply “not derived demand.” These are the user’s alternative formulations.'],
 ['Four categories','The four categories in PaidToBeHuman are optional communication aids. They are not required fields or a checklist of necessary conditions in this assessment.'],
 ['Original work read','PaidToBeHuman/source/Paid_to_Be_Human_Conversation_Export.md, original User passages at lines 184, 223, 265, 298 and 337: machine-saturated supply, firefighters, nonfungible people, human presence, provenance and personally accountable contractors. User passages were distinguished from Assistant replies.'],
 ['Additional source context','Reviewed the finished “derived vs essential demand” transcript, current edited course Lessons 02, 04, 06, 07 and 08, and the white-paper sections on definitions and measurement. Supporting synthesis was not substituted for the current user’s definition.'],
 ['Root source path','C:\\PaidToBeHuman\\source\\Paid_to_Be_Human_Conversation_Export.md'],
 ['Supporting source paths','C:\\PaidToBeHuman\\reference\\transcripts\\derived vs essential demand.md; C:\\PaidToBeHuman\\course\\Lesson 02/04/06/07/08 (current .docx files); C:\\PaidToBeHuman\\reference\\background\\Derived vs Essential Demand v01 (text).md'],
 ['Estimated quantity','For each occupation, the estimated proportion of present labor demand that is essential rather than derived under the user’s definition. The coefficients are judgments about the occupation’s mixture, not measured percentages of customers, fees or task minutes.'],
 ['Low / median / high','Low and high are plausible judgment bounds. Median is the subjective central or best estimate, not an empirical sample median, fitted 50th percentile or arithmetic midpoint. The bounds are not statistical confidence intervals.'],
 ['Derived-demand complements','Derived low = 100% − essential high. Derived median = 100% − essential median. Derived high = 100% − essential low. Opposite endpoints therefore sum to 100%.'],
 ['Employment weighting','Multiply each occupation’s 2025 jobs by its essential coefficient; sum those job-weighted equivalents and divide by the sum of scored jobs. Do not average occupation percentages equally. These are job-weighted equivalents, not full-time equivalents or identifiable jobs that will survive.'],
 ['Judgment and independence','Every detailed occupation has an individually authored estimate and explanation. No target aggregate, automatic positive floor or four-category allocation was imposed. Earlier author audit numbers and Claude estimates were not used to calibrate these scores.'],
 ['Mixed occupations','The explanation identifies what customers or institutions obtain and the human-specific reason, if any. A coefficient may be zero at one, two or all three bounds. Broad “all other” groups receive wider bounds where their mix is unclear.'],
 ['Assumptions behind aggregation','The current occupational mix is held fixed. Future changes in prices, service volume, entry, technology, hours, preferences and institutions are not modeled. Summing every low or high coefficient gives aligned scenarios, not a probability interval for the economy.'],
 ['Employment baseline','2025 National Employment Matrix jobs from BLS Table 1.2, 2025–35 projections release of August 27, 2026. Only the historical 2025 employment column is used. This is not a real-time November 2026 jobs count.'],
 ['Universe','The matrix includes nonfarm wage-and-salary employment, self-employment, agricultural and private-household employment. It is a civilian job/employment measure, not unique employed people. This workbook covers every published Table 1.2 “Line item” occupation.'],
 ['Rows included','831 distinct Line item occupations in 22 major groups. Summary rows are retained on BLS source for audit but are excluded from occupation weighting to avoid double counting. Leading indentation is removed from titles; codes and numeric values are unchanged.'],
 ['Coverage reconciliation','The detailed rows sum to 170,267,200 jobs; the published total is 170,280,800. The 13,600 difference is not assigned to an invented occupation. BLS notes that subcategories can differ because of rounding and excluded occupational data. The headline estimates use covered employment.'],
 ['Unpublished occupations','For example, legislators are not a published Line item in this table. They are not given an invented employment weight. The workbook is complete for the supplied table, rather than an assertion that every conceivable occupation has a separately published row.'],
 ['Raw employment precision','BLS reports employment in thousands with one decimal place. Conversion to jobs multiplies by 1,000, preserving source precision of 100 jobs. Computed equivalents are rounded for display only.'],
 ['Supplied input',meta.source_path+'; worksheet '+meta.source_sheet],
 ['Input SHA-256',meta.sha256],
 ['Official BLS workbook',meta.source_url],
 ['Fresh official table verification',meta.official_html+' — retrieved November 16, 2026. All 831 detailed codes, titles and 2025 employment values exactly match the supplied workbook.'],
 ['BLS definitions','https://www.bls.gov/emp/documentation/definitions.htm'],
 ['BLS methodology','https://www.bls.gov/emp/documentation/projections-methods.htm'],
 ['BLS release','https://www.bls.gov/news.release/ecopro.htm'],
 ['Occupation scope references','https://www.bls.gov/soc/2018/soc_2018_definitions.pdf; occupation-specific OOH links are preserved from the supplied workbook in Occupations. Scope descriptions inform duties; they do not supply the essential-demand percentages.'],
 ['Review corrections','Checked all codes, employment values, ordered bounds and complements. Clarified postsecondary teaching-assistant duties and fitness-studies teaching, and made the childcare-versus-preschool assumption explicit. Numerical estimates were not changed to match a target result.'],
 ['Using the workbook','Filter Occupations by code, title, group or estimate. Pale yellow cells are editable essential-demand estimates. Derived shares, job-weighted equivalents and Summary recalculate from those inputs. Keep low ≤ median ≤ high, each between 0% and 100%.']
];
body(method,`A1:B${notes.length+4}`);method.tabColor='#8E9DAF';
title(method,'A2','Method and sources');
method.getRange('A4:B4').values=[['Topic','Definition, assumption or source']];header(method,'A4:B4');
method.getRange(`A5:B${notes.length+4}`).values=notes;
method.getRange('B8').setNumberFormat('yyyy-mm-dd hh:mm:ss "UTC"');
method.getRange(`A1:A${notes.length+4}`).format.columnWidth=31;
method.getRange(`B1:B${notes.length+4}`).format.columnWidth=128;
method.getRange(`A5:B${notes.length+4}`).format.wrapText=true;
method.getRange(`A5:B${notes.length+4}`).format.verticalAlignment='center';
for(let i=0;i<notes.length;i++){
 const row=i+5;const lines=Math.ceil(notes[i][1].length/132);
 method.getRange(`A${row}:B${row}`).format.rowHeight=Math.max(36,lines*15+14);
 if(i%2===0)method.getRange(`A${row}:B${row}`).format.fill='#F3F6FA';
}
method.freezePanes.freezeRows(4);

// Recalculation checks a reversible input edit, then restores the exact original estimate.
const original=occ.getRange('E6').values[0][0];
const originalSum=summary.getRange('C10').values[0][0];
occ.getRange('E6').values=[[original+0.01]];
const changed=summary.getRange('C10').values[0][0];
const expectedDelta=rows[0].employment_2025*0.01;
if(Math.abs(changed-originalSum-expectedDelta)>0.001)throw Error('Input-change recalculation failed');
occ.getRange('E6').values=[[original]];
w.recalculate();
const core=summary.getRange('B8:D10').values;
for(let i=0;i<3;i++){
 const key=['low','median','high'][i];
 if(Math.abs(core[0][i]-expected.essential_shares[key])>1e-10)throw Error('Summary share mismatch '+key);
 if(Math.abs(core[2][i]-expected.essential_equivalents[key])>0.001)throw Error('Summary equivalents mismatch '+key);
}
const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:40},maxChars:3000});
await fs.writeFile(path.join(here,'formula_error_scan.ndjson'),errors.ndjson);
const inspection=await w.inspect({kind:'table',range:'Summary!A7:D18',include:'values,formulas',tableMaxRows:12,tableMaxCols:4,maxChars:7000});
await fs.writeFile(path.join(here,'summary_inspection.ndjson'),inspection.ndjson);
await fs.writeFile(path.join(here,'calculation_checks.json'),JSON.stringify({recalculationDelta:changed-originalSum,expectedDelta,summary:core,zeroCounts:summary.getRange('B47:B48').values},null,2));
// Render compact representative views from every sheet, plus the long explanation column.
for(const[sheetName,range,file]of [
 ['Occupations','A249:J250','preview_numeric_precision.png']
]){
 const blob=await w.render({sheetName,range,scale:1.5,format:'png'});
 await fs.writeFile(path.join(here,file),new Uint8Array(await blob.arrayBuffer()));
}
const output=await SpreadsheetFile.exportXlsx(w);
const target=path.join(outdir,run.filename);
await output.save(target);
await fs.writeFile(path.join(here,'output_path.txt'),target);
console.log(JSON.stringify({output:target,rows:rows.length,summary:core,recalculationDelta:changed-originalSum,errorScan:errors.ndjson},null,2));
