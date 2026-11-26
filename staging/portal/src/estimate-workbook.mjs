// Export only the current contract's demo estimates, not invoices or actual work.
export function createEstimateWorkbook(ExcelJS,contract,items,columns,summarize){
 const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('Stima lavorazioni'),summary=wb.addWorksheet('Riepilogo');
 const fields=columns.filter(([field])=>!['dataEsecuzione','oraEsecuzione','operatoreNome','latitudine','longitudine','stato','note'].includes(field));
 ws.addRow([...fields.map(c=>c[1]),'Verifica economica']);
 for(const [i,item] of items.entries()){
  const values={...item,numeroProgressivoRiga:i+1};
  ws.addRow([...fields.map(([f])=>values[f]??''),item.economicStatus==='OK'?'OK':'DA VERIFICARE']);
 }
 ws.views=[{state:'frozen',ySplit:1}];ws.autoFilter={from:{row:1,column:1},to:{row:1,column:fields.length+1}};
 fields.forEach(([f],i)=>{
  const col=ws.getColumn(i+1);col.width=['denominazione','indirizzo','tipologiaLavorazione'].includes(f)?34:22;
  if(['prezzoBase','prezzoRibassato','totale'].includes(f))col.numFmt='€ #,##0.00';
 });
 ws.getColumn(fields.length+1).width=24;
 for(const sheet of [ws,summary]){sheet.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};sheet.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF166534'}};}
 const report=summarize(items);
 summary.addRows([['STIMA DIMOSTRATIVA — NON CONSUNTIVO',''],['Commessa',String(contract.nome||'')],['Codice commessa',String(contract.codice||'')],['Lavorazioni',report.rows],['Valorizzate',report.verified],['Da verificare',report.missing],['Importo stimato parziale',report.estimated],['Nota','Solo importi valorizzati; non attesta lavorazioni eseguite, IVA o fatturazione.']]);
 summary.getColumn(1).width=45;summary.getColumn(2).width=85;summary.getCell('B7').numFmt='€ #,##0.00';
 return wb;
}
