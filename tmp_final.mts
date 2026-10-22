import { readFileSync } from "fs";
import { read, utils } from "xlsx";
const p="/home/ubuntu/attachments/02ea6744-0293-4d68-947c-0d46196aee06/ITS_AND_RED_BUS_PNR_-_Copy_1.xlsx";
const wb=read(readFileSync(p));
console.log(wb.SheetNames);
const ws=wb.Sheets['FINAL'];
const rows=utils.sheet_to_json(ws,{header:1,raw:false}) as any[][];
console.log("rows:",rows.length);
rows.slice(0,30).forEach((r,i)=>{ if(r.some(v=>v!==undefined&&v!=="")) console.log(i+1,JSON.stringify(r)); });
