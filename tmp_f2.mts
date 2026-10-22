import { readFileSync } from "fs";
import { read, utils } from "xlsx";
const p="/home/ubuntu/attachments/02ea6744-0293-4d68-947c-0d46196aee06/ITS_AND_RED_BUS_PNR_-_Copy_1.xlsx";
const wb=read(readFileSync(p),{cellFormula:true});
const ws=wb.Sheets['FINAL'];
const hdr=utils.sheet_to_json(ws,{header:1,raw:false}) as any[][][0];
for(const k in ws){
  if(k[0]==='!')continue;
  const m=k.match(/^([A-Z]+)(\d+)$/); if(!m)continue;
  if(parseInt(m[2])!==2)continue;
  const c=ws[k];
  console.log(k, hdr ? '' : '', c.f?`FORMULA: ${c.f}`:'', 'v=',c.w);
}
