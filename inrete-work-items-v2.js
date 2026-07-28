/* Prezzi e migrazione INRETE v2. Funzioni pure condivise da Web e runtime Android. */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.InreteWorkItemsV2 = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const normalizePriceCode = value => String(value ?? "").trim().toLocaleUpperCase("it-IT");
  const normalizeStatus = value => {
    const status = String(value ?? "DA FARE").trim().toLocaleUpperCase("it-IT").replace(/_/g, " ");
    return status === "DONE" || status === "COMPLETATO" ? "FATTO" : status;
  };
  const parseNumber = value => {
    if (value === "" || value == null) return null;
    let text = String(value).trim().replace(/\s/g, "");
    if (text.includes(",")) text = text.replace(/\./g, "").replace(",", ".");
    const number = Number(text);
    return Number.isFinite(number) ? number : null;
  };
  const splitLegacyValues = value => [...new Set(String(value ?? "").split(/[|;,\r\n]+/).map(v => v.trim()).filter(Boolean))];
  const toDate = value => {
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
    if (value && typeof value.toDate === "function") return toDate(value.toDate());
    if (value && Number.isFinite(Number(value.seconds))) return new Date(Number(value.seconds) * 1000);
    if (typeof value === "string" || typeof value === "number") {
      const parsed = new Date(value);
      return Number.isNaN(parsed.getTime()) ? null : parsed;
    }
    return null;
  };
  const executionFields = (value, fallbackDate = "", fallbackTime = "") => {
    const date = toDate(value);
    if (!date) return {dataEsecuzione:String(fallbackDate || ""), oraEsecuzione:String(fallbackTime || "")};
    const part = number => String(number).padStart(2, "0");
    return {dataEsecuzione:`${date.getFullYear()}-${part(date.getMonth()+1)}-${part(date.getDate())}`, oraEsecuzione:`${part(date.getHours())}:${part(date.getMinutes())}`};
  };
  const adaptLegacyPlantToWorkItems = plant => {
    const codes = splitLegacyValues(plant?.codicePrezzo || plant?.voceRiferimento || plant?.codiceVocePrezzo);
    if (!codes.length) codes.push("");
    const descriptions = splitLegacyValues(plant?.tipologiaIntervento || plant?.tipologiaLavorazione || plant?.tipologiaImpianto);
    const quantities = splitLegacyValues(plant?.quantitaPerCodice);
    const completed = executionFields(plant?.doneAt, plant?.dataEsecuzione || plant?.dataFatto, plant?.oraEsecuzione || plant?.oraFatto);
    return codes.map((code, index) => ({
      ...plant,
      id: codes.length === 1 ? plant.id : `${plant.id}__${normalizePriceCode(code).replace(/[^A-Z0-9]+/g,"_") || index + 1}`,
      impiantoId: plant.id,
      legacy: true,
      codiceVocePrezzo: code,
      quantita: quantities.length === codes.length ? parseNumber(quantities[index]) : (codes.length === 1 ? (plant.quantita ?? plant.areaMq ?? "") : null),
      tipologiaLavorazione: descriptions.length === codes.length ? descriptions[index] : (plant.tipologiaIntervento || plant.tipologiaLavorazione || plant.tipologiaImpianto || ""),
      dataEsecuzione: completed.dataEsecuzione,
      oraEsecuzione: completed.oraEsecuzione,
      operatoreNome: plant.operatore || plant.doneBy || "",
      stato: plant.done ? "FATTO" : (plant.stato || "DA FARE"),
      quantityVerificationRequired: codes.length > 1 && quantities.length !== codes.length
    }));
  };
  const buildPriceMap = prices => new Map((prices || []).map(item => [normalizePriceCode(item.codiceVoce ?? item.codiceVocePrezzo ?? item.codice), item]));
  const resolvePriceItem = (pricesOrMap, code) => (pricesOrMap instanceof Map ? pricesOrMap : buildPriceMap(pricesOrMap)).get(normalizePriceCode(code)) || null;
  const effectiveDiscount = (general, specific) => parseNumber(specific) ?? parseNumber(general) ?? 0;
  const calculateDiscountedPrice = (base, general, specific) => {
    const amount = parseNumber(base), discount = effectiveDiscount(general, specific);
    return amount == null || discount < 0 ? null : amount * (1 - discount);
  };
  const calculateWorkItemTotal = (workItem, priceItem, generalDiscount) => {
    if (!normalizePriceCode(workItem?.codiceVocePrezzo) || !priceItem) return null;
    const discounted = calculateDiscountedPrice(priceItem.prezzoBase, generalDiscount, priceItem.percentualeRibasso);
    if (discounted == null) return null;
    if (String(priceItem.unitaMisura || "").trim().toUpperCase() === "AC") return discounted;
    const quantity = parseNumber(workItem.quantita);
    return quantity == null ? null : quantity * discounted;
  };
  const enrichWorkItem = (workItem, pricesOrMap, generalDiscount) => {
    const priceItem = resolvePriceItem(pricesOrMap, workItem.codiceVocePrezzo);
    if (!priceItem) return {...workItem, unitaMisura:null, prezzoBase:null, percentualeRibasso:null, prezzoRibassato:null, totale:null, priceListItemId:null, priceListLinkStatus:"MISSING", economicStatus:"DA_VERIFICARE"};
    const percentualeRibasso = effectiveDiscount(generalDiscount, priceItem.percentualeRibasso);
    const prezzoRibassato = calculateDiscountedPrice(priceItem.prezzoBase, generalDiscount, priceItem.percentualeRibasso);
    const totale = calculateWorkItemTotal(workItem, priceItem, generalDiscount);
    return {...workItem, codiceVocePrezzo:String(workItem.codiceVocePrezzo).trim(), unitaMisura:priceItem.unitaMisura || "", prezzoBase:parseNumber(priceItem.prezzoBase), percentualeRibasso, prezzoRibassato, totale, priceListItemId:priceItem.id || null, priceListLinkStatus:"LINKED", economicStatus:totale == null ? "DA_VERIFICARE" : "OK"};
  };
  const calculateCompletedSubtotal = items => (items || []).reduce((sum, item) => normalizeStatus(item.stato) === "FATTO" && Number.isFinite(parseNumber(item.totale)) ? sum + parseNumber(item.totale) : sum, 0);
  const isInreteCommessa = commessa => ["nome","codice","categoria","tipo","commessaPadre","parentName"].some(key => normalizePriceCode(commessa?.[key]).includes("INRETE"));
  const derivePlantStatus = items => {
    if (items.some(item => item.economicStatus === "DA_VERIFICARE")) return "DA VERIFICARE";
    const states=items.map(item=>normalizeStatus(item.stato));
    if (states.includes("IN LAVORAZIONE")) return "IN LAVORAZIONE";
    if (states.every(s=>s==="FATTO")) return "FATTO";
    if (states.every(s=>s==="DA FARE")) return "DA FARE";
    if (states.includes("FATTO")) return "PARZIALMENTE FATTO";
    return states[0] || "DA FARE";
  };

  async function migrateInreteCommesseToWorkItemsV2(options) {
    const {db, collectionName="globalCommesse", currentUser, isAdmin, operatorName=""} = options || {};
    if (!db || !currentUser || !isAdmin) throw new Error("Migrazione riservata agli amministratori.");
    const commesseSnap=await db.collection(collectionName).get();
    const report={commesseAnalizzate:0,impiantiAnalizzati:0,impiantiMigrati:0,lavorazioniCreate:0,impiantiMultiCodice:0,lavorazioniFatteCreate:0,codiciRiconosciuti:0,codiciMancanti:0,quantitaDaVerificare:0,ignoratiGiaMigrati:0,errori:[],dettagli:[]};
    for (const commDoc of commesseSnap.docs) {
      const commessa={id:commDoc.id,...commDoc.data()}; if(!isInreteCommessa(commessa)) continue;
      report.commesseAnalizzate++; if(Number(commessa.inreteMigrationVersion)>=2){report.ignoratiGiaMigrati++;continue;}
      const ref=commDoc.ref, log=db.collection("inreteMigrationAudit").doc(`${commDoc.id}_${Date.now()}`), startedAt=new Date();
      const baseLog={migrationId:log.id,migrationType:"INRETE_WORK_ITEMS_V2",commessaId:commDoc.id,commessaNome:commessa.nome||"",plantsScanned:0,plantsMigrated:0,workItemsCreated:0,completedWorkItemsCreated:0,missingPriceCodes:0,quantityWarnings:0,startedAt,completedAt:null,requestedByUid:currentUser.uid,requestedByName:operatorName,status:"RUNNING",errorMessage:""};
      await log.set(baseLog);
      try {
        const [legacySnap, pricesSnap, existingSnap]=await Promise.all([ref.collection("impianti").get(),ref.collection("prezziario").get(),ref.collection("lavorazioni").get()]);
        const prices=pricesSnap.docs.map(d=>({id:d.id,...d.data()})), priceMap=buildPriceMap(prices), existing=new Map(existingSnap.docs.map(d=>[d.data().migrationSourceId,d]).filter(([sourceId])=>sourceId));
        let created=0, completed=0, missing=0, warnings=0, migrated=0;
        for(const legacyDoc of legacySnap.docs){
          const legacy={id:legacyDoc.id,...legacyDoc.data()}; report.impiantiAnalizzati++; const codes=splitLegacyValues(legacy.codicePrezzo||legacy.voceRiferimento||legacy.codiceVocePrezzo);
          if(!codes.length) codes.push(""); if(codes.length>1) report.impiantiMultiCodice++;
          const descriptions=splitLegacyValues(legacy.tipologiaIntervento||legacy.tipologiaLavorazione); const descriptionsMatch=descriptions.length===codes.length;
          const quantities=splitLegacyValues(legacy.quantitaPerCodice); const ambiguous=codes.length>1 && quantities.length!==codes.length && legacy.quantita!=null;
          const plantRef=ref.collection("impiantiFisici").doc(legacyDoc.id);
          await plantRef.set({commessaId:commDoc.id,numeroProgressivoImpianto:legacy.numeroProgressivo||null,idSap:legacy.idSap||legacy.idSAP||"",denominazione:legacy.denominazione||legacy.nome||"",comune:legacy.comune||"",indirizzo:legacy.indirizzo||legacy.via||"",latitudine:legacy.latitudine??legacy.gpsY??null,longitudine:legacy.longitudine??legacy.gpsX??null,distretto:legacy.distretto||"",note:legacy.note||legacy.noteImpianto||"",migrationSourceId:legacyDoc.id},{merge:true});
          const generated=[];
          for(let i=0;i<codes.length;i++){
            const code=codes[i], sourceId=`${legacyDoc.id}::${normalizePriceCode(code)||"NO_CODE"}`;
            if(existing.has(sourceId)){
              const existingDoc=existing.get(sourceId), existingData=existingDoc.data(), completed=executionFields(legacy.doneAt,legacy.dataEsecuzione||legacy.dataFatto,legacy.oraEsecuzione||legacy.oraFatto);
              if(normalizeStatus(legacy.done?"FATTO":legacy.stato)==="FATTO"&&(!existingData.dataEsecuzione||!existingData.oraEsecuzione))await existingDoc.ref.set({...completed,doneAt:legacy.doneAt||null,operatoreNome:existingData.operatoreNome||legacy.operatore||legacy.doneBy||""},{merge:true});
              report.ignoratiGiaMigrati++;continue;
            }
            const priceItem=resolvePriceItem(priceMap,code), historicalStatus=normalizeStatus(legacy.done?"FATTO":legacy.stato), quantity=quantities.length===codes.length?parseNumber(quantities[i]):(!ambiguous?parseNumber(legacy.quantita??legacy.areaMq):null);
            const completed=executionFields(legacy.doneAt,legacy.dataEsecuzione||legacy.dataFatto,legacy.oraEsecuzione||legacy.oraFatto);
            let item=enrichWorkItem({commessaId:commDoc.id,impiantoId:legacyDoc.id,migrationSourceId:sourceId,codiceVocePrezzo:code,quantita:quantity,tipologiaLavorazione:descriptionsMatch?descriptions[i]:(legacy.tipologiaIntervento||legacy.tipologiaLavorazione||""),stato:historicalStatus,...completed,operatoreNome:legacy.operatore||legacy.doneBy||"",operatoreUid:legacy.operatoreUid||legacy.doneByUid||""},priceMap,commessa.percentualeRibassoGenerale);
            if(ambiguous && String(item.unitaMisura).toUpperCase()!=="AC") item={...item,quantita:null,totale:null,economicStatus:"DA_VERIFICARE",quantityVerificationRequired:true};
            if(!descriptionsMatch && descriptions.length>1) item={...item,economicStatus:"DA_VERIFICARE",descriptionVerificationRequired:true};
            const workRef=ref.collection("lavorazioni").doc(`${legacyDoc.id}__${normalizePriceCode(code).replace(/[^A-Z0-9]+/g,"_")||"NO_CODE"}`); await workRef.set(item,{merge:true}); existing.set(sourceId,{ref:workRef,data:()=>item}); generated.push(item); created++; if(historicalStatus==="FATTO")completed++; if(!priceItem)missing++;
          }
          if(ambiguous){warnings++;report.dettagli.push(`Impianto ${legacy.denominazione||legacy.nome||legacyDoc.id}: presenti più codici prezzo ma una sola quantità. Verificare le quantità delle singole lavorazioni.`);}
          if(normalizeStatus(legacy.done?"FATTO":legacy.stato)==="FATTO"&&!legacy.dataEsecuzione&&!legacy.dataFatto)report.dettagli.push(`${legacy.denominazione||legacyDoc.id}: Data storica non disponibile.`);
          await plantRef.set({stato:derivePlantStatus(generated.length?generated:[{stato:legacy.stato}])},{merge:true}); migrated++;
        }
        Object.assign(report,{impiantiMigrati:report.impiantiMigrati+migrated,lavorazioniCreate:report.lavorazioniCreate+created,lavorazioniFatteCreate:report.lavorazioniFatteCreate+completed,codiciRiconosciuti:report.codiciRiconosciuti+created-missing,codiciMancanti:report.codiciMancanti+missing,quantitaDaVerificare:report.quantitaDaVerificare+warnings});
        await ref.set({excelModelVersion:2,priceListVersion:2,workItemsModelVersion:2,inreteMigrationVersion:2,inreteMigratedAt:new Date(),inreteMigratedBy:currentUser.uid},{merge:true});
        await log.set({plantsScanned:legacySnap.size,plantsMigrated:migrated,workItemsCreated:created,completedWorkItemsCreated:completed,missingPriceCodes:missing,quantityWarnings:warnings,completedAt:new Date(),status:"COMPLETED",errorMessage:""},{merge:true});
      } catch(error) { report.errori.push(`${commessa.nome||commDoc.id}: ${error.message}`); await log.set({completedAt:new Date(),status:"FAILED",errorMessage:error.message||String(error)},{merge:true}); }
    }
    return report;
  }
  return {normalizePriceCode,normalizeStatus,parseNumber,splitLegacyValues,toDate,executionFields,adaptLegacyPlantToWorkItems,buildPriceMap,resolvePriceItem,calculateDiscountedPrice,calculateWorkItemTotal,enrichWorkItem,calculateCompletedSubtotal,isInreteCommessa,derivePlantStatus,migrateInreteCommesseToWorkItemsV2};
});
