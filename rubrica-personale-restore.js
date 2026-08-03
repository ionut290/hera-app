(() => {
  'use strict';

  const RUN_KEY = 'personale-safe-restore-20260803-v1';
  if (window.__vargaPersonaleSafeRestore20260803 || localStorage.getItem(RUN_KEY) === 'done') return;
  window.__vargaPersonaleSafeRestore20260803 = true;
  window.__personaleDestructiveClearDisabled = true;

  const ROWS = [["0QeSqsoisQhb22XnR17Q","ASIF","MOHAMMAD","","","","","","NO","",""],["0ZiPZfxtuBaN7s3u60nW","BARBARA","STEGANI","","","","","","NO","",""],["0olsMHgOtR7iOaOGEmdv","MASSIMO","MURAN","","","","","","NO","",""],["1MaQpvbCHjTMeyENVlcg","FLAVIO","GIORGIO","","","","","","NO","",""],["2f3QkVQygTNpU6uiWdAE","RAPHAEL","OKOSUN","","","","","","NO","",""],["2mL017VsX2eGppsu6CNC","PAOLO","CHIAPELLI","","","","","","NO","",""],["2xXty9i0gqF0MMDc6oM4","ARSLAN","FARDOOS","","","","","","NO","",""],["3o4zlQFFcr1i3NQ44hkJ","LORENZO","BOTTITTA","","","","","","NO","",""],["4RjcF47UBLUsciHBBsy5","SALVATORE","PIRRITANO","","","","","","NO","",""],["57O7f9EdkpvGRWSKoTyn","PIERO","NANNI","3297667658","","","","","NO","",""],["5DZ7XUFFYYwKNLUqIQi0","SALOME","EGBONA","","","","","","NO","",""],["5jMbOCnXsgulwyO78r5J","CHIRIAC DRAGOS","NICOLAIE","","","","","","NO","",""],["5qvhXiunYcap4D2RUJ5y","SEIED","HEIDAR HOSEINI","","","","","","NO","",""],["5sjfuHTHsX2QC60lZYiJ","IBRAHIM","ZORJANI","","","","","","NO","",""],["65Lt6e1Csu5ahhboHpqh","LORENZO","CALICIOTTI","","","","","","NO","",""],["69WBInJdYspNMR3UOtBT","IMRAN","KHAN","","","","","","NO","",""],["6Bbj6X4n5fDgBc7hLmya","CHRISTABEL","SUNDAY","","","","","","NO","",""],["6dSO239gANvsrS06cZJd","MICHELA","ZAMBELLI","","","","","","NO","",""],["6gLOcIKrpBMr1Pu1it8Z","HENRRY","CONTRERAS","","","","","","NO","",""],["72g9TxYfYUF6ZmgFkWG1","HASSAN","SHABBIR","","","","","","NO","",""],["85d9lhiZAnXRY7sohaqK","SILVANO","MARANGON","","","","","","NO","",""],["8HKkvjoa9auYbO4DP9a3","ROLANDO","ROMAN LUYA","","","","","","NO","",""],["8L93CC4i6xUltkGAmQAC","ANTONIO","BARONE","","baroneantonio059@gmail.com","baroneantonio059@gmail.com","","","NO","xaevQispOClLjqb7J6BZ,77RxprlNHFvdLzt48IEm,A7nR3tjyuKo7jCyacWGq,iziGhrLtZosDYt7OnhQA,nuVKqqZAOTCWLLsirnwB,VItH5xiyGXHekHjfImmq,q1MXY9MpZF3F1hU4vhBl,HGSbSp7CQ5W08rS1RQ0R,ekgoFPYGMeYQiGPUJDOI,584PNXNPzH1TNqgYzeYL,Rt83Xhm1WqVSgpb6BbYJ,oiDsqNevc97l9OsZ6Oy1,cKotT7roXPjJG3JiaGZj,vg3CoYHV8gYZxmcIpEF0,WS0dkqR2tszh2CQeRCNa,dz2OJibaIJuMic55toQP",""],["97AHhh8l19DXYMsfCiJP","ABDELMOULA","ERRACHIDI","","","","","","NO","",""],["AMGPL947jISQQd4Izbpc","MASSIMO","FERRARI","","","","","","NO","",""],["AgmL9Ql9KU87ncgeGkIb","FEDERICO","GHERARDI","","","","","","NO","",""],["AnoVeXS9KKbEu7eIUVIv","FEDERICO","ARGINELLI","","","","","","NO","",""],["BFxViHYR06qawWLTXWyd","EMILIAN","CRISTIAN GHITIU","","","","","","NO","",""],["BTmnrhyKwvyYRZ7PI4Td","MICHEL","DAMANDE","","","","","","NO","",""],["BeTHsvAhDjHfrgHpaK3A","FEDERICO","TARTARINI","","","","","","NO","",""],["BlcCd92WknlwDHqLf9hf","SIMONA","PAULA BUDACA","","","","","","NO","",""],["DNIrDOAHQFpDQ9HiqcKH","KHAN","MUHAMMAD","","","","","","NO","",""],["DPP51rXVdKUvKzqL9IUc","MATTEO","PIZZIRANI","","","","","","NO","",""],["DS6XMCzooPjQd9rhbpUZ","GIOVANNI BRANDO","DI","","","","","","NO","",""],["EFwHZvZPfy8UdB9n4KWt","BROLLO MILO","DAL","","miloxdalxbrollox@gmail.com","miloxdalxbrollox@gmail.com","","","NO","xaevQispOClLjqb7J6BZ,77RxprlNHFvdLzt48IEm,A7nR3tjyuKo7jCyacWGq,iziGhrLtZosDYt7OnhQA,nuVKqqZAOTCWLLsirnwB,VItH5xiyGXHekHjfImmq,q1MXY9MpZF3F1hU4vhBl,HGSbSp7CQ5W08rS1RQ0R,ekgoFPYGMeYQiGPUJDOI,584PNXNPzH1TNqgYzeYL,Rt83Xhm1WqVSgpb6BbYJ,oiDsqNevc97l9OsZ6Oy1,cKotT7roXPjJG3JiaGZj,vg3CoYHV8gYZxmcIpEF0,WS0dkqR2tszh2CQeRCNa,dz2OJibaIJuMic55toQP",""],["Eer44RYh79UAT8BYXc15","EMANUELE","TALIGNANI","","","","","","NO","",""],["Ew3DjJfhU6CheGp3QHbi","MAMADOU","KONE","","","","","","NO","",""],["F2aYDY7zWWMoiBdMzpak","VASCO","SAVORRI","","","","","","NO","",""],["F4SEbWFUefsutbJ7gGES","SALVATORE","VITALE","","","","","","NO","",""],["FA1WIXgvTS0qVxBroXga","VLADIMIR","MEREACRE","","","","","","NO","",""],["G0E2poZD4GpElBYH0DTr","GIULIO","UGOLINI","","","","","","NO","",""],["GFuP9rEwOfp0ew22Pbp3","ERIKA","STARINIERI","","","","","","NO","",""],["HwZqjLcjgpMmuLTNYm4j","GABI-CORNEL","STOICHINA","","","","","","NO","",""],["JEYaTCsEWAOADIzJBcZw","MATTEO","FERRETTI","","","","","","NO","",""],["JWo6kovx2SuFWRWZcAWy","FAITH","OMORU","","","","","","NO","",""],["JhHEzIfV7BwE0ixR50Uk","MANUEL","GUERNELLI","","","","","","NO","",""],["KFIgF73Ww6icVD8G2geX","FRANCESCO","OSTI","","","","","","NO","",""],["KYyj2S8Zcqy3wuB70NM6","ANDREA","ZARRI","","andreazarri@yahoo.it","andreazarri@yahoo.it","2wreh7ij3CcjZayO3jaE9DIT3la2","andreazarri@yahoo.it","SÌ","xaevQispOClLjqb7J6BZ,77RxprlNHFvdLzt48IEm,A7nR3tjyuKo7jCyacWGq,iziGhrLtZosDYt7OnhQA,nuVKqqZAOTCWLLsirnwB,VItH5xiyGXHekHjfImmq,q1MXY9MpZF3F1hU4vhBl,HGSbSp7CQ5W08rS1RQ0R,ekgoFPYGMeYQiGPUJDOI,584PNXNPzH1TNqgYzeYL,Rt83Xhm1WqVSgpb6BbYJ,oiDsqNevc97l9OsZ6Oy1,cKotT7roXPjJG3JiaGZj,vg3CoYHV8gYZxmcIpEF0,WS0dkqR2tszh2CQeRCNa,dz2OJibaIJuMic55toQP",""],["Kxb08n6TVEqEtG5i8KgC","MOHAMED ALI MAHMOUD","FATHI","","","","","","NO","",""],["L7xI2iLnUbD3flzjYlsD","LUCIO","CORONA","","","","","","NO","",""],["L85wCKFIwxTR580QCghK","MAME","MOR SAMB","","","","","","NO","",""],["LEoeHPAez0ZO1lWjO2hw","EMMANUEL","NUAOBASI","","","","","","NO","",""],["M1g2J16q58QlyK7dn0EO","BILAL","MUHAMMAD","","","","","","NO","",""],["N1QOmedFLPULeLqLmpBJ","LAZAR RELU","LUCIAN","","","","","","NO","",""],["NBp2ZcQsfTzFcHOqZUiT","IONEL","VARGA","","ionut29019@gmail.com","ionut29019@gmail.com","KzVjYtKABxbglt3YyDRbOkZypBg2","ionut29019@gmail.com","SÌ","q1MXY9MpZF3F1hU4vhBl,xaevQispOClLjqb7J6BZ,77RxprlNHFvdLzt48IEm,A7nR3tjyuKo7jCyacWGq,iziGhrLtZosDYt7OnhQA,nuVKqqZAOTCWLLsirnwB,VItH5xiyGXHekHjfImmq,HGSbSp7CQ5W08rS1RQ0R,ekgoFPYGMeYQiGPUJDOI,Rt83Xhm1WqVSgpb6BbYJ,oiDsqNevc97l9OsZ6Oy1,cKotT7roXPjJG3JiaGZj,vg3CoYHV8gYZxmcIpEF0,dz2OJibaIJuMic55toQP","GOOGLE"],["NKMgNVtDe6ZDwNegobXf","HELEN","SALAMI","","","","","","NO","",""],["NRNEicPo9toclB93W8vP","LORENZO","SACCHI","","","","","","NO","",""],["NTXk7kbnWTfUUNTWvxV0","ANDREA","SUPPINI","","","","","","NO","",""],["O5Y1RoFgSgwrYCtSMBpB","GIOVANNI","POZZI","","","","","","NO","",""],["OftlrlFKdBX4T77YxzWk","MARIO","SCOLLO","","","","","","NO","",""],["PWX1f2ocjO3uvvOG7KGZ","AZIZ","AAZIZ","","","","","","NO","",""],["PYMVKxaQndUtJ18yEHNA","TARAS","KRYSKIV","","","","","","NO","",""],["PuJ1kGX88AsI7D4PqzRN","RACHID","BOUHCHICH","","","","","","NO","",""],["QLEaFwXUWadSuU4Uu3Aq","FABIO","MAGGIORI","","","","","","NO","",""],["QLlUtN6B7hcEGeXBAcyx","VALENTIN","DRAGULEAN","","","","","","NO","",""],["Qeb0M5xecfKcAGmtt18u","MICHAEL","ROY MENDOZA MATOS","","","","","","NO","",""],["RFS6fBNxKB6V9aEdY8FI","MARZIA","MINELLI","","","","","","NO","",""],["SxKB5HQ3KlMeD9FZIAAe","AKRAM","MUHAMMAD","","","","","","NO","",""],["TSO3Z0B6JTZhfxTsfLII","VALENTINA","GERARDI","","","","","","NO","",""],["TbXBVbls69wJwcN7kvvX","IMURANA","ALI","","","","","","NO","",""],["UKUkfC9KV4ukY9EZrxXf","ION","JOITA","","","","","","NO","",""],["URQzgV0kM7Tf9WRQ2AVR","PAOLO","RUSSO","","","","","","NO","",""],["V70SiH8vRlTsBTuFrmtk","MODOU","NDAW","","","","","","NO","",""],["VLWtr9FzwEzudoa39Yn4","ALEXANDRU","LAZAR","","","","","","NO","",""],["VaBlSe8Nujmoh7vOaZH4","CORRADO","CENERINI","","","","","","NO","",""],["Vf7Dnga7tawnFmQG6ZFd","LUIS","JUSTO CHAMBI ARAPA","","","","","","NO","",""],["VtCAg20NSDsRbWDmI9mM","MUBASHAR","JAVED","","","","","","NO","",""],["WHqjoKz0zPsNHOlL3Vd8","CHERKAOUI","ED DAHRI","","","","","","NO","",""],["WifEQdnvZHWYX4GrvUgx","ALESSANDRO","FRANCESCHELLI","","","","","","NO","",""],["X4dnfkQTL14du2YrgcS3","OSATO","OJEGA","","","","","","NO","",""],["XLFxNavsSZoJQ3EnKZaG","CHAKER","OUCHARI","","","","","","NO","",""],["Xmiypexqwyc9qzXYNwbf","RICKY","LAVEZZI","","","","","","NO","",""],["Y7ljiDAhLOqovi4XScAv","LUIGI","ANTONIO MARTINO","","","","","","NO","",""],["YDqPinEoVeZmiHauTCcr","TERESA","FINI","","","","","","NO","",""],["YqbzLAVRdG9t33ygHneZ","AMABILE","BOLDRIN","","","","","","NO","",""],["YzdGNdvIS00MeHzK93xg","IGOR","COVALIOV","","","","","","NO","",""],["Z11AvFO0klUg3Mtgh4pq","DARIEL","VELAZQUEZ FERNANDEZ","","","","","","NO","",""],["ZHe6HAyGjm5gHIEUsg5q","BUJA DANIEL","G.","","","","","","NO","",""],["Zx4oxHevRavBLBizdrCJ","MAKSYM","ZARUBKO","","","","","","NO","",""],["aQcTa2wkgcEeiW4AwrJj","CRISTIAN","SITZIA","","","","","","NO","",""],["aTfPJs8kAedRRQXQu3jO","JOEY","RITZ JAURIGUE","","","","","","NO","",""],["bCMImgJrTnYDu9os6bwX","BILAWAL","IQBAL","","","","","","NO","",""],["blZBkkbIqrFsv55Mfnj4","MUHAMMAD","ADNAN IQBAL","","","","","","NO","",""],["cQ1lhQ5MNcZVXExDmYPH","GEORGETA","GERDAN","","","","","","NO","",""],["cmCAvM3MqiLwD4QkU3hb","GIUSEPPE","ARTUSO","","","","","","NO","",""],["dHVlCBE8gpk6PDq5jo1t","MANUEL","PARZANESE","","","","","","NO","",""],["dJ8CQUZA5blaUVldybow","LORENZA","MINELLI","","","","","","NO","",""],["dUVE39jITNoCXfUJaxXy","GIORGIO","AGOSTINO DE GIORGIO","","","","","","NO","",""],["dWowKvgnNpGpgOyPQYbT","SERAFIM","COJOCARU","","","","","","NO","",""],["drrY66tCNLwISO3BnYqJ","PIETRO","GALVAGNI","","","","","","NO","",""],["eJh7xzA0Gd0XLwmWnX7S","ANTONIO","IACOVIELLO","","","","","","NO","",""],["et710uj9BwTN5qPV83nQ","ROBERT","LEGUIA HUAYANA","","","","","","NO","",""],["fFeWPV0hjiLSuQbB4wpU","FABIO","GOLINELLI","","","","","","NO","",""],["gQNDhyQMsfnRJaF7WOPW","TIZIANA","BALBONI","","","","","","NO","",""],["gQcGqg5pp3Eu6zrQ21a0","ABDELKERIM","LAMSSAOUI","","","","","","NO","",""],["gZziSsuUclOT7ScMNTTg","JOHNSON","YEBOAH","","sqadrahera@gmail.com","sqadrahera@gmail.com","","","NO","xaevQispOClLjqb7J6BZ,77RxprlNHFvdLzt48IEm,A7nR3tjyuKo7jCyacWGq,iziGhrLtZosDYt7OnhQA,nuVKqqZAOTCWLLsirnwB,VItH5xiyGXHekHjfImmq,q1MXY9MpZF3F1hU4vhBl,HGSbSp7CQ5W08rS1RQ0R,ekgoFPYGMeYQiGPUJDOI,584PNXNPzH1TNqgYzeYL,Rt83Xhm1WqVSgpb6BbYJ,oiDsqNevc97l9OsZ6Oy1,cKotT7roXPjJG3JiaGZj,vg3CoYHV8gYZxmcIpEF0,WS0dkqR2tszh2CQeRCNa,dz2OJibaIJuMic55toQP",""],["grgwJvQtXKaW9KGs6fE9","MOUSTAPHA","SYLLA","","","","","","NO","",""],["j5oaUJtvSGWIme5t8VjA","CRISTIAN-ADRIAN","POP","","","","","","NO","",""],["jpnqJT6KO75bN0paRyBJ","ALBERT","AMOAH FRIMPONG","","","","","","NO","",""],["jq9t2L4PaMKmChESuh5Z","DAVIDE","MARTINENGO","","","","","","NO","",""],["kV4T7UK4DNojPTHzPICB","ANGELA","VENTURI GRANDI","","","","","","NO","",""],["kx5XdmvOXiCejQ56FA8g","CESAR","OMAR PEREZ PONCE","","","","","","NO","",""],["lArqz1VUmj5DzhnTUGCu","ERNEST","ODURO","","","","","","NO","",""],["lBbVdOFoiAvaU9rRtl5r","IOAN","CIPRIAN FLOREA","","","","","","NO","",""],["mONAE97pi99hbEGrEPgz","FATOS","TURUKU","","","","","","NO","",""],["nsvbL6tOm3R02jlHZAqt","ANTONINO","PROVENZANO","","","","","","NO","",""],["oN6GbsqtlPDjJCMt3DXo","ELISA","LANDI","","","","","","NO","",""],["oovch1bmNdngyZzS1FLh","MEVLAN","ZORJANI","","","","","","NO","",""],["owYjAYUvIRROkICqP1AW","ALESSIA","MONARI","","","","","","NO","",""],["pEmEWhr9QMFyD4akmy9U","FEDERICO","RUSSO","","","","","","NO","",""],["pHzkZoWEEZ4a6DqX3buV","ALEXANDR","TEAN","","","","","","NO","",""],["qWGpc51GgI7DAFsYWRcB","MUHAMMAD","AJMAL","","","","","","NO","",""],["rTcIGDvKVLR4UdoaHVuh","GIUSEPPE","LEVATO","","","","","","NO","",""],["rngEpvrqnnXoMimQK2L2","ADRIANO","MENNA","","","","","","NO","",""],["ssYui5aU7HPSJ86ttqzr","PRINCE","AMA OKORIE","","","","","","NO","",""],["uwTp0eaM2VwJBnugBBSB","PIETRO","PIRRITANO","","","","","","NO","",""],["vEXmHtRt87VBbjhGSMC1","LUCA","MELONI","","","","","","NO","",""],["vOUAS5GP5uVoUlqJPwdG","BELLO","NOUHADINE EYONULAGBA","","","","","","NO","",""],["vXFeIb8DKf02kp2rWt0T","DANIEL","VIESTI","","","","","","NO","",""],["vbFYi77BMPtkHVWHsl4w","FABIO","ADAMES RAMIREZ","","","","","","NO","",""],["vdF0UQmPr9sCUycxSNSn","HAFEDH","BEJAOUI","","","","","","NO","",""],["vxiidZg4XQgrwtfFWBvu","ALESSANDRO","BASCIU","","","","","","NO","",""],["wA9XITZZlKsUWH1AlRsc","VITO","LIMA","","","","","","NO","",""],["wtBl9PFtQtjkqUNnch5V","LORENZO","PIZZIRANI","","","","","","NO","",""],["x1HwzMb5sn33A5Sa0aUd","ELMOSTAFA","MELIANI","","","","","","NO","",""],["x1NVcOrhB407xhuMdUCf","MASSIMO","MASOTTI","","","","","","NO","",""],["xCObRQnJeZWNDnx3d80L","ANDREA","FRIGHI","","","","","","NO","",""],["HaLWKEpHSGA5QMm2vnAK","MIRKO","SANDONI","","","","","","NO","",""],["0hY3nArD3eqRJ8DrDGFQ","STEFANO","MECAGNI","","","","","","NO","",""],["0vOGiakA6ie0PYyHT9AV","S.","MACAGNI","","","","","","NO","",""],["QVqdZLXnxNd7kyDBGJWb","NACCAH BENITO PIETRO","E","","pietroenaccah88@gmail.com","pietroenaccah88@gmail.com","Nzh1BRk3xxeDoFNvXutuiH3RUnk1","pietroenaccah88@gmail.com","SÌ","xaevQispOClLjqb7J6BZ,77RxprlNHFvdLzt48IEm,A7nR3tjyuKo7jCyacWGq,iziGhrLtZosDYt7OnhQA,nuVKqqZAOTCWLLsirnwB,VItH5xiyGXHekHjfImmq,q1MXY9MpZF3F1hU4vhBl,HGSbSp7CQ5W08rS1RQ0R,ekgoFPYGMeYQiGPUJDOI,584PNXNPzH1TNqgYzeYL,Rt83Xhm1WqVSgpb6BbYJ,oiDsqNevc97l9OsZ6Oy1,cKotT7roXPjJG3JiaGZj,vg3CoYHV8gYZxmcIpEF0,WS0dkqR2tszh2CQeRCNa,dz2OJibaIJuMic55toQP","GOOGLE"],["xr27lNG1hmmXJxX9CaSC","DISTACO ENERGREEN","PERINI","","","","","","NO","",""],["6oCutbHg0mDScApUW8BG","Muddsir","Rehman","3512916599","mudasarrehman582@gimail.com","mudasarrehman582@gimail.com","","","NO","",""]];
  const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const splitCommesse = (value) => String(value || '').split(',').map((item) => item.trim()).filter(Boolean);

  const RECORDS = ROWS.map((row) => ({
    idOperatore: row[0],
    nome: row[1],
    cognome: row[2],
    telefono: row[3] || '',
    email: row[4] || '',
    stato: 'ATTIVO',
    attivo: true,
    emailAccessoApp: row[5] || '',
    linkedUserId: row[6] || '',
    linkedUserEmail: row[7] || '',
    profiloCollegato: row[8] || 'NO',
    commesseAbilitate: splitCommesse(row[9]),
    modalitaAggiornamentoCommesse: 'NON_MODIFICARE',
    mantieniAbilitazioniEsistenti: true,
    fonteFotoProfilo: row[10] || '',
    restoredFrom: 'Matrice_Personale_STATO_ATTIVO_2026-08-03'
  }));

  async function restoreMissingPersonale() {
    for (let attempt = 0; attempt < 120; attempt += 1) {
      const firestore = window.firebase?.firestore?.();
      const user = window.firebase?.auth?.()?.currentUser;
      const isManager = typeof window.canManageData === 'function' && window.canManageData();

      if (!firestore || !user || !isManager) {
        await wait(250);
        continue;
      }

      const collectionName = typeof window.getPersonaleCollectionName === 'function'
        ? window.getPersonaleCollectionName()
        : 'personale';
      const collection = firestore.collection(collectionName);
      const snapshot = await collection.get();
      const existingIds = new Set(snapshot.docs.map((doc) => doc.id));
      const missing = RECORDS.filter((record) => record.idOperatore && !existingIds.has(record.idOperatore));

      if (!missing.length) {
        localStorage.setItem(RUN_KEY, 'done');
        console.info(`Ripristino Personale: tutti i ${RECORDS.length} record risultano presenti.`);
        return;
      }

      for (let start = 0; start < missing.length; start += 400) {
        const batch = firestore.batch();
        missing.slice(start, start + 400).forEach((record) => {
          const ref = collection.doc(record.idOperatore);
          batch.set(ref, {
            ...record,
            createdAt: window.firebase.firestore.FieldValue.serverTimestamp(),
            updatedAt: window.firebase.firestore.FieldValue.serverTimestamp()
          }, { merge: false });
        });
        await batch.commit();
      }

      localStorage.setItem(RUN_KEY, 'done');
      window.dispatchEvent(new CustomEvent('personale-safe-restore-complete', {
        detail: { restored: missing.length, total: RECORDS.length, collectionName }
      }));
      console.info(`Ripristino Personale completato: ${missing.length} record ricreati su ${RECORDS.length}.`);

      try {
        if (typeof window.subscribePersonale === 'function') window.subscribePersonale();
        else if (typeof window.renderPersonale === 'function') window.renderPersonale();
      } catch (error) {
        console.warn('Personale ripristinato; aggiornamento vista non riuscito:', error);
      }
      return;
    }

    throw new Error('Firestore o permessi amministratore non disponibili.');
  }

  restoreMissingPersonale().catch((error) => {
    window.__vargaPersonaleSafeRestore20260803 = false;
    console.error('Ripristino sicuro Personale non riuscito:', error);
  });
})();
