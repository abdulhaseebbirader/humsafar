import { readFileSync } from "fs";
import { parseCsv, reconcileCsv } from "./src/lib/reconcile";
const v = parseCsv(readFileSync("/tmp/up/red_bus.csv","utf8"));
// print headers + one row per section
let section="";
for(const row of v){
  const t=(row[0]??"").trim();
  if(/-(Payment|Cancellation|Bus Cancelled|Exceptional Refund)$/i.test(t)){section=t;console.log("\n=== SECTION:",t);continue}
  if(/^TIN$/i.test(t)){console.log("HDR:",row.map(c=>c.trim()).join(" | "));continue}
  if(section && row[1]&&/\d/.test(row[1])&&/\d/.test(row[2]??"")){console.log("ROW:",row.map(c=>c.trim()).join(" | "));section="";}
}
