const fields = Object.freeze(['tipoRecord','denominazione','codice','comune','localita','via','civico','indirizzo','specieAlbero','specieSiepe','note','numeroAlbero','diametroCm','altezzaAlberoM','lunghezzaM','altezzaSiepeM','larghezzaM','superficieMq','tipologiaIntervento','lavorazioniRichieste','gpsX','gpsY','gpsAccuracyM','gpsDetectedAt','coordinateSource','commessaId','commessaNome','commessaCodice','cap','provincia','regione','paese','statoVegetativo','source']);
export function projectVerdeRecords(snapshot) {
  if (!Array.isArray(snapshot?.docs)) throw new Error('Invalid record snapshot');
  return snapshot.docs.map(doc => {
    if (typeof doc.id !== 'string' || !doc.id || doc.id.includes('/')) throw new Error('Invalid record ID');
    const data = doc.data();
    if (!data || !['cantiere','albero','siepe'].includes(data.tipoRecord)) throw new Error('Invalid record category');
    const record = {id:doc.id};
    for (const field of fields) {
      const value = data[field];
      if (typeof value === 'string' || (typeof value === 'number' && Number.isFinite(value))) record[field] = value;
    }
    return Object.freeze(record);
  });
}
