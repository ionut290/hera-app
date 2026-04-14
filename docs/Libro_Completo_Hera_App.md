# Libro completo di utilizzo — Hera App

Versione guida: 14 aprile 2026  
Applicazione analizzata: repository `hera-app` (web app HTML/CSS/JS con Firebase)

---

## 1) Introduzione

Hera App è una piattaforma operativa per cantieri, pensata per smartphone e tablet, che integra in una sola interfaccia:

- gestione commesse,
- gestione impianti e interventi,
- gestione squadre/personale/mezzi,
- chat operativa,
- esportazione verso Google Sheets,
- archiviazione documenti/media su Google Drive,
- funzioni meteo, GPS, servizi territoriali, segnalazioni e gestione ore.

L'app è una **PWA** (Progressive Web App) e può essere anche pubblicata come app Android tramite Capacitor.

---

## 2) Requisiti e accesso

### 2.1 Requisiti minimi

- Browser moderno (Chrome, Edge, Safari recente, Firefox recente).
- Connessione Internet stabile per sync cloud.
- Account Google per autenticazione utente.

### 2.2 Avvio locale (sviluppo/test)

1. Aprire terminale nella cartella del progetto.
2. Avviare server statico:

```bash
python3 -m http.server 8080
```

3. Aprire `http://localhost:8080`.

### 2.3 Login e account

- Pulsante principale: **Login con Google**.
- Pulsante **Cambia account** per forzare la scelta account Google.
- Pulsante **Logout** per uscita sessione.

---

## 3) Architettura funzionale (visione semplice)

L'app è organizzata in grandi aree operative:

1. **Home Dashboard**: commesse, stato utente, meteo operativo, azioni rapide.
2. **Pagina Impianti**: elenco impianti, ricerca, filtri “Da fare/Fatti”, mappa e navigazione.
3. **Area Gestione (Management)**: creazione e amministrazione dati principali.
4. **Funzioni verticali**: chat, ore, documenti personali, segnalazioni, servizi personali, FAQ “Come si fa?”.
5. **Integrazioni cloud**: Firebase Auth/Firestore, Google Drive, Google Sheets.

---

## 4) Guida completa per area (uso pratico)

## 4.1 Home Dashboard

Nella Home trovi:

- riepilogo utenti attivi,
- ultima azione impianto,
- prossima azione consigliata,
- meteo operativo con alert rischio,
- blocco utente/notifiche,
- elenco commesse,
- blocco squadre per commessa.

### Flusso consigliato inizio giornata

1. Effettua login.
2. Verifica meteo operativo.
3. Seleziona commessa.
4. Apri composizione squadre e verifica mezzo/personale.
5. Passa a Impianti e avvia il giro (ordine vicino→lontano).

---

## 4.2 Commesse

Funzioni disponibili:

- creazione commessa,
- selezione commessa attiva,
- gestione elenco commesse lato admin,
- eventuale associazione Google Sheet dedicato.

### Import impianti in una commessa

Puoi importare da:

- Excel (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`),
- CSV, ODS, FODS,
- HTML/TXT tabellare,
- URL Google Sheet.

L’app normalizza i campi e unisce righe duplicate dello stesso impianto.

---

## 4.3 Impianti (core operativo)

Quando entri nella pagina impianti:

- visualizzi mappa e lista,
- gli impianti sono ordinati per distanza,
- puoi cercare per denominazione/comune/indirizzo/codice,
- puoi passare tra tab **Da fare** e **Fatti**,
- puoi esportare gli impianti fatti (quando disponibile).

### Azioni tipiche su impianto

- apri navigazione,
- apri WhatsApp contestuale,
- marca stato di avanzamento,
- apri editor impianto,
- apri modulo report impianto.

### Stato “Fatto”

Quando un impianto viene marcato “Fatto”:

- viene registrata data/ora,
- viene registrato operatore,
- se configurato, viene preparata/sincronizzata riga su Google Sheet della commessa.

---

## 4.4 Mappa, GPS e geolocalizzazione

Funzioni principali:

- geolocalizzazione utente,
- marker impianti con stile variabile,
- popup impianto,
- focus elemento lista da mappa,
- modalità mappa a schermo intero,
- calcolo distanze con formula haversine.

Se il GPS non è disponibile, l’app mostra stato e fallback senza bloccare tutta l’operatività.

---

## 4.5 Squadre

Area dedicata a:

- composizione squadra per commessa,
- associazione personale/mezzi,
- data squadra (automatica o manuale),
- suggerimenti e compilazione assistita,
- invio proposta squadre su WhatsApp.

### Buona pratica

- Compila prima personale e mezzi anagrafici.
- Poi genera la squadra con righe multiple personale+mezzo.
- Verifica suggerimenti e invia proposta.

---

## 4.6 Personale

Funzioni:

- aggiunta manuale nominativi,
- import da file (quando previsto nella UI),
- visualizzazione e cancellazione record,
- uso dei nominativi nelle squadre e nella gestione ore.

---

## 4.7 Mezzi

Funzioni:

- anagrafica mezzo con campi tecnici (N-ID, marca, modello, portata, massa, alimentazione),
- import da file,
- ricerca/normalizzazione N-ID,
- riuso automatico dati mezzi nelle squadre.

---

## 4.8 Chat operatori

Caratteristiche:

- chat completa con invio testo,
- messaggi broadcast o destinatario specifico,
- invio media (foto/video),
- invio vocali,
- contatore non letti,
- gestione lettura messaggi.

### Media su Drive

Quando il bridge Drive è attivo, i media possono essere caricati su Drive con link condivisibile.

---

## 4.9 Documenti personali

Area riservata con:

- upload file,
- eventuale acquisizione da camera,
- note e nome documento,
- preset rapidi (es. PIN/tessera),
- salvataggio su Drive (in base configurazione).

---

## 4.10 Gestione ore

Workflow:

1. Seleziona data.
2. Aggiungi una o più commesse nel report ore.
3. Inserisci righe operatore + ore.
4. Controlla riepilogo (duplicati/consistenza).
5. Finalizza report.
6. Consulta vista mensile/tabella.
7. Esporta tabella ore (se prevista dalla UI).

Funzioni avanzate:

- totali per operatore,
- totali operatore+commessa,
- filtri mensili,
- storico report salvati,
- eventuali livelli approvativi.

---

## 4.11 Segnalazioni

Modulo completo con:

- preposto,
- data/ora,
- cantiere,
- descrizione,
- presa visione,
- firme (tecnico/preposto),
- condivisione via WhatsApp o email.

---

## 4.12 Servizi personali

Funzione geolocalizzata per trovare servizi nelle vicinanze con:

- mappa dedicata,
- lista servizi,
- categorie filtrabili,
- raggio ricerca impostabile,
- schede dettagliate (indirizzo, info utili, stato buoni pasto se disponibile).

---

## 4.13 Meteo operativo

La dashboard meteo evidenzia rischio operativo su condizioni come:

- pioggia,
- neve,
- nebbia.

È disponibile un dettaglio esteso in modal con informazioni aggiuntive.

---

## 4.14 Notifiche PWA

Sono presenti:

- pulsante **Attiva notifiche**,
- pulsante **Test notifica**,
- gestione stato permessi.

Il Service Worker gestisce:

- `push`,
- `notificationclick`,
- `sync` (background check).

Con chiave VAPID pubblica configurata, l’app può usare push remote reali.

---

## 4.15 Google Drive + Google Sheets

### Drive bridge centralizzato

- Un admin collega una volta Drive.
- Token e cartelle vengono salvati in Firestore.
- Gli altri utenti autenticati usano il bridge condiviso.

### Cartelle create automaticamente

- `Hera App - Dati`
- `Hera App - Dati/Chat Media`
- `Hera App - Dati/Report Impianti`

### Export su Sheet

Per ogni commessa:

- si può usare Sheet pre-esistente (ID/link), oppure
- se assente viene creato automaticamente,
- al “Fatto” viene aggiunta riga con dati impianto + timestamp + operatore.

---

## 4.16 Sezione Global

La modalità **Global** è separata dalle commesse standard:

- commesse global dedicate,
- import file globale,
- import da Google Sheet globale,
- ricerca e mappa global.

Nessuna interferenza diretta con il flusso standard principale.

---

## 4.17 Gestione utenti e permessi

Area admin per:

- aggiunta/rimozione utenti amministratori,
- visualizzazione utenti,
- gestione azioni negate/consentite,
- elenco app esterne,
- richieste GPS,
- riepiloghi stato utente.

---

## 5) Procedure operative complete (SOP)

## SOP A — Nuova commessa da zero

1. Login admin.
2. Apri pannello “Aggiungi commesse”.
3. Crea nuova commessa.
4. Seleziona commessa target.
5. Importa file impianti o Sheet URL.
6. Verifica lista impianti e mappa.
7. Comunica al team la disponibilità.

## SOP B — Giornata operatore

1. Login.
2. Seleziona commessa.
3. Apri impianti.
4. Segui ordine di prossimità.
5. Usa navigazione e azioni contestuali.
6. Segna impianti completati.
7. A fine giornata verifica tab “Fatti”.

## SOP C — Report ore giornaliero

1. Apri gestione ore.
2. Inserisci data.
3. Aggiungi commesse trattate.
4. Compila ore per operatore.
5. Verifica riepilogo.
6. Finalizza e salva.

## SOP D — Segnalazione sicurezza/qualità

1. Apri Segnalazioni.
2. Compila campi obbligatori.
3. Acquisisci firme.
4. Condividi via canale richiesto.
5. Archivia riferimento in team.

---

## 6) Troubleshooting (problemi comuni)

### 6.1 Login non riuscito

- Verifica popup consentiti nel browser.
- Verifica account Google autorizzato.
- Riprova con “Cambia account”.

### 6.2 Notifiche non attive

- Controlla permesso notifiche browser.
- Verifica supporto Service Worker.
- Se push remote non arriva, controlla chiave VAPID.

### 6.3 Mappa non mostra posizione

- Concedi permesso geolocalizzazione.
- Verifica GPS del dispositivo attivo.
- Controlla connessione dati.

### 6.4 Export Sheet non immediato

- Verifica connessione.
- L’app usa code/retry locali per alcune mutazioni.
- Controlla eventuali errori token Drive.

### 6.5 Media chat/documenti non caricati

- Controlla connessione.
- Controlla che Drive bridge sia collegato.
- Controlla dimensione file rispetto limiti.

---

## 7) Sicurezza e governo del dato

Linee guida consigliate:

- usare account nominativi (non condivisi),
- limitare privilegi admin al minimo necessario,
- verificare periodicamente autorizzazioni e token,
- mantenere FAQ/help center aggiornati dopo nuove feature,
- applicare policy di retention per chat/media secondo norme aziendali.

---

## 8) Pubblicazione Android (senza rompere il web)

L’app web resta la fonte principale. Android è un wrapper con Capacitor.

Comandi tipici:

```bash
npm install
npm run android:add
npm run android:sync
npm run android:open
```

In Android Studio:

- Generate Signed Bundle/APK,
- scegliere Android App Bundle (AAB),
- firmare e pubblicare su Play Console.

---

## 9) Dizionario rapido dei pannelli menu

- Aggiungi commesse
- Composizione squadre
- Personale
- Mezzi
- Gestione utenti
- Global
- Informazioni utili
- Documenti personali
- Servizi personali
- Gestione ore
- Segnalazioni
- Come si fa?

---

## 10) Appendice A — Inventario completo funzioni JavaScript

Di seguito l’inventario delle funzioni dichiarate in `app.js` al momento della generazione di questo libro.

1. `resolvePushPublicVapidKey()`
2. `isAutoNotificationEnabled()`
3. `setAutoNotificationEnabled()`
4. `syncNotificationAutoPreferenceFromProfile()`
5. `toggleUserDetailsPanel()`
6. `updateNotificationUi()`
7. `urlBase64ToUint8Array()`
8. `subscribeGlobalNotifications()`
9. `stopGlobalNotificationsSubscription()`
10. `weatherCodeLabel()`
11. `updateAdminControls()`
12. `openSideMenu()`
13. `closeSideMenu()`
14. `openManagementPanel()`
15. `closeManagementPanel()`
16. `toggleMapFullscreen()`
17. `toggleMapCssFullscreen()`
18. `syncMapFullscreenUiState()`
19. `applyRoute()`
20. `openImpiantiPage()`
21. `closeImpiantiPage()`
22. `closeFuelPage()`
23. `openPersonalServicesPage()`
24. `closePersonalServicesPage()`
25. `setCurrentWorkflowStep()`
26. `getWorkflowSteps()`
27. `renderNextActionCard()`
28. `getCurrentImpiantoNextAction()`
29. `impiantoNextActionLabel()`
30. `impiantoNextActionIcon()`
31. `buildInlineActionButton()`
32. `renderImpiantoNextActionUI()`
33. `toggleImpiantoNextActionHighlight()`
34. `registerImpiantoSessionAction()`
35. `openSegnalazioniPage()`
36. `closeSegnalazioniPage()`
37. `openHowtoPage()`
38. `closeHowtoPage()`
39. `buildHowtoFaqItems()`
40. `openPrivateDocsPage()`
41. `closePrivateDocsPage()`
42. `initHoursPage()`
43. `openHoursPage()`
44. `closeHoursPage()`
45. `renderSavedHoursReports()`
46. `openHoursViewModal()`
47. `closeHoursViewModal()`
48. `getMonthMeta()`
49. `resolveHoursStatsMonth()`
50. `ensureHoursViewModalOpen()`
51. `renderHoursMonthlyTable()`
52. `renderHoursOperatoriOptions()`
53. `renderHoursCommessaSelectOptions()`
54. `renderHoursTableCommessaOptions()`
55. `addHoursOperatoreRow()`
56. `addHoursCommessaBlock()`
57. `collectHoursEntries()`
58. `normalizeHoursOperatorName()`
59. `findDuplicateHoursInDraft()`
60. `renderHoursSummary()`
61. `setHoursFinalizeLocked()`
62. `unlockHoursFinalizeButton()`
63. `renderHowtoFaq()`
64. `prefillSegnalazioneDateTime()`
65. `syncSegnalazioneFirmaPreposto()`
66. `getSegnalazioneData()`
67. `validateSegnalazioneData()`
68. `loadPendingSheetExports()`
69. `savePendingSheetExports()`
70. `startSheetRetryLoop()`
71. `queuePendingSheetExport()`
72. `getStoredDriveToken()`
73. `trackLocalSheetMutation()`
74. `hasRecentLocalSheetMutation()`
75. `scheduleCommessaSheetSync()`
76. `parseGoogleSheetId()`
77. `isAndroidWebViewRuntime()`
78. `loginWithGoogle()`
79. `persistDriveAccessToken()`
80. `logout()`
81. `subscribeDriveBridge()`
82. `stopDriveBridgeSubscription()`
83. `subscribeCommesse()`
84. `stopCommesseSubscription()`
85. `subscribeGlobalCommesse()`
86. `stopGlobalCommesseSubscription()`
87. `onGlobalCommessaSelectionChanged()`
88. `subscribeGlobalImpianti()`
89. `stopGlobalImpiantiSubscription()`
90. `onGlobalExcelSelected()`
91. `onGlobalImpiantoSearchInput()`
92. `renderGlobalImpianti()`
93. `renderGlobalMap()`
94. `subscribeResources()`
95. `stopResourcesSubscription()`
96. `subscribePrivateDocs()`
97. `stopPrivateDocsSubscription()`
98. `applyPrivateDocPreset()`
99. `getPrivateDocsDriveToken()`
100. `renderPrivateDocsList()`
101. `renderResourcesList()`
102. `renderResourceManageFilters()`
103. `resourceTypeLabel()`
104. `getResourcesByCommessa()`
105. `renderResourceButtonsForCommessa()`
106. `openCommessaResourceWindow()`
107. `closeCommessaResourceViewer()`
108. `renderCommessaResourceViewer()`
109. `resourceTypeLongLabel()`
110. `updateResourceFormByType()`
111. `sanitizePhone()`
112. `openPhoneOnWhatsApp()`
113. `openDocumentLink()`
114. `downloadVCard()`
115. `selectCommessa()`
116. `updateCommessaContextUI()`
117. `updateCommessaButtonsActive()`
118. `onCommessaTargetChanged()`
119. `getTargetCommessaId()`
120. `getTargetCommessaName()`
121. `subscribeImpianti()`
122. `stopImpiantiSubscription()`
123. `onExcelSelected()`
124. `normalizeRow()`
125. `mergeRowsByImpianto()`
126. `combineImpiantiForView()`
127. `buildImpiantoKey()`
128. `mergeMultiValue()`
129. `normalizeKeys()`
130. `getValue()`
131. `getValueWithMatchedKey()`
132. `mergeExtraFields()`
133. `rowsFromWorkbookBuffer()`
134. `buildGoogleSheetCsvUrl()`
135. `parseCoordinate()`
136. `classifyTipoManutenzione()`
137. `firestoreDateToMillis()`
138. `formatDoneDateTime()`
139. `renderHeaderActivitySummary()`
140. `splitCodes()`
141. `buildRowsForEachCodicePrezzo()`
142. `hasOrdinario()`
143. `hasStraordinario()`
144. `onImpiantoSearchInput()`
145. `setImpiantiViewMode()`
146. `matchesImpiantoSearch()`
147. `renderImpianti()`
148. `openImpiantoEditor()`
149. `closeImpiantoEditor()`
150. `updateConnectivityStatus()`
151. `createButton()`
152. `createActionIconButton()`
153. `setUsedActionButtonState()`
154. `isActionUsed()`
155. `markActionAsUsed()`
156. `clearActionUsed()`
157. `normalizeMezzoNId()`
158. `mergeMezzoData()`
159. `buildMezzoPatch()`
160. `findExistingMezzoByNId()`
161. `normalizeHeaderKey()`
162. `extractPersonnelNamesFromRawRows()`
163. `subscribePersonale()`
164. `subscribeMezzi()`
165. `subscribeSquadre()`
166. `stopPersonaleSubscription()`
167. `stopMezziSubscription()`
168. `stopSquadreSubscription()`
169. `renderSimpleList()`
170. `renderMezziList()`
171. `autofillSquadraForm()`
172. `addSquadraRow()`
173. `resolveSuggestionValue()`
174. `parseMultiEntryValue()`
175. `addMultiEntryInput()`
176. `renumberSquadraRows()`
177. `readSquadraRows()`
178. `getLegacySquadreRows()`
179. `setSquadraRowsFromData()`
180. `getDateKeyFromLocalDate()`
181. `getAutomaticSquadreDateKey()`
182. `getActiveSquadreDateKey()`
183. `onSquadreFilterDateChange()`
184. `clearManualSquadreFilterDate()`
185. `renderSquadre()`
186. `renderMezziButtonsMarkup()`
187. `updateSquadraHintFromSources()`
188. `updateSuggestionLists()`
189. `getPersonaleDisplayName()`
190. `canTriggerImpiantoWhatsApp()`
191. `triggerImpiantoWhatsAppAction()`
192. `openWhatsApp()`
193. `openImpiantoReportModal()`
194. `closeImpiantoReportModal()`
195. `getCurrentPositionOnce()`
196. `openSquadraWhatsApp()`
197. `getSquadrePackageEntries()`
198. `getCommessaAccentColor()`
199. `renderMap()`
200. `buildImpiantoMapPopup()`
201. `focusImpiantoInList()`
202. `cssEscapeValue()`
203. `getMarkerClass()`
204. `updateImpiantoLocalState()`
205. `getImpiantoDocIds()`
206. `canManageData()`
207. `normalizeEmail()`
208. `clearMap()`
209. `getMezzoByLabel()`
210. `toggleFuelMezzoDetails()`
211. `renderFuelMezzoDetails()`
212. `detectFuelBrand()`
213. `ensureFuelMap()`
214. `renderFuelStations()`
215. `createFuelMarkerIcon()`
216. `getFuelMarkerClass()`
217. `onPersonalServiceCategoryClick()`
218. `normalizePersonalServices()`
219. `defaultPersonalServiceName()`
220. `ensurePersonalServicesMap()`
221. `clearPersonalServicesMap()`
222. `renderPersonalServicesMap()`
223. `createPersonalServiceMarkerIcon()`
224. `getPersonalServiceMarkerClass()`
225. `renderPersonalServicesList()`
226. `renderLunchGroupedList()`
227. `buildPersonalServiceRow()`
228. `selectPersonalService()`
229. `buildExpandedPersonalServiceDetails()`
230. `renderExtendedPersonalServiceDetails()`
231. `formatDetailFieldLabel()`
232. `formatAddress()`
233. `isMealVoucherAccepted()`
234. `formatMealVoucherStatus()`
235. `getSelectedPersonalServicesRadius()`
236. `initGeolocation()`
237. `buildImpiantoEventLocalKey()`
238. `hasLocalImpiantoEvent()`
239. `markLocalImpiantoEvent()`
240. `evaluateImpiantoProximityAlerts()`
241. `renderWeatherDetails()`
242. `openWeatherModal()`
243. `closeWeatherModal()`
244. `distanceFromUser()`
245. `haversine()`
246. `formatDistance()`
247. `getTrafficIntensityByHour()`
248. `getDistanceIntensityOffset()`
249. `getRouteVarianceOffset()`
250. `normalizeTrafficIntensity()`
251. `estimateTravelMeta()`
252. `escapeHTML()`
253. `subscribeChat()`
254. `isChatMessageFresh()`
255. `startChatRetentionLoop()`
256. `stopChatRetentionLoop()`
257. `subscribeAdminUsers()`
258. `stopAdminUsersSubscription()`
259. `subscribeUsers()`
260. `stopUsersSubscription()`
261. `canApproveHoursLevel1()`
262. `renderHoursApprovalRequests()`
263. `startPresenceHeartbeat()`
264. `stopPresenceHeartbeat()`
265. `subscribeGpsRequests()`
266. `stopGpsRequestsSubscription()`
267. `renderGpsRequests()`
268. `getDeniedActionsForCurrentUser()`
269. `isImpiantoActionDenied()`
270. `actionLabel()`
271. `renderUserPermissionList()`
272. `renderExternalApps()`
273. `renderChatRecipients()`
274. `renderCommesseManagementList()`
275. `renderAdminUsers()`
276. `stopChatSubscription()`
277. `countUnreadMessages()`
278. `canViewMessage()`
279. `markChatAsRead()`
280. `canConfirmHoursFromChat()`
281. `setChatHoursActionButtonsState()`
282. `openChatModal()`
283. `closeChatModal()`
284. `isOwnMessage()`
285. `enforceMediaSize()`
286. `resetDriveState()`
287. `updateDriveStatus()`
288. `extractGoogleAccessToken()`
289. `normalizeFaqData()`
290. `escapeHtml()`
291. `toIsoDate()`
292. `renderFaqHelpCenter()`
293. `getCommessaSheetHeaders()`
294. `buildSheetRowsFromDoneImpianti()`

Totale funzioni dichiarate: **294**

---

## 11) Appendice B — Integrazioni tecniche principali

- Firebase Authentication
- Firestore con persistenza offline (best effort)
- Firebase Messaging (se disponibile)
- Service Worker per PWA/push/sync
- Google Drive API (bridge centralizzato)
- Google Sheets API (export avanzamenti)
- Leaflet (mappa)
- SheetJS/XLSX (import file tabellari)

---

## 12) Nota finale

Questo libro è “completo” rispetto alle funzionalità presenti nel codice corrente. Se vuoi, nel prossimo step posso generarti anche:

1. versione **manuale utente semplificato** (solo operatore),
2. versione **manuale admin** separato,
3. versione **manuale formazione con checklist giornaliera**,
4. versione con **copertina aziendale e indice cliccabile**.
