(() => {
  "use strict";

  const byId = (id) => document.getElementById(id);
  const form = byId("private-docs-form");
  const nameInput = byId("private-docs-name");
  const noteInput = byId("private-docs-note");
  const fileInput = byId("private-docs-file");
  const cameraInput = byId("private-docs-camera");
  const saveButton = byId("private-docs-save-btn");
  const feedback = byId("private-docs-feedback");
  const tesseraButton = byId("private-docs-preset-tessera-btn");
  const flow = byId("identity-capture-flow");
  const genericFields = [byId("private-docs-generic-file-field"), byId("private-docs-generic-camera-field")];
  const video = byId("identity-camera-video");
  const dedicatedInput = byId("identity-camera-input");
  const cropDialog = byId("identity-crop-dialog");
  const cropCanvas = byId("identity-crop-canvas");
  const cropHandles = byId("identity-crop-handles");
  const cropPreview = byId("identity-crop-preview");
  if (!form || !window.firebase?.firestore) return;

  const MAX_PDF_BYTES = 500000;
  const IDENTITY_RATIO = 85.6 / 53.98;
  let identityPreset = false;
  let stream = null;
  let sourceCanvas = null;
  let processedIdentity = null;
  let corners = [];
  let activeHandle = -1;
  let previousBodyOverflow = "";

  const showFeedback = (message, type = "") => {
    if (!feedback) return;
    feedback.textContent = message;
    feedback.classList.toggle("private-docs-feedback-success", type === "success");
    feedback.classList.toggle("private-docs-feedback-error", type === "error");
  };
  const readAsDataUrl = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Non riesco a leggere il file selezionato."));
    reader.readAsDataURL(file);
  });
  const loadImage = (url) => new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("La foto selezionata non è valida."));
    image.src = url;
  });

  // Reads JPEG orientation before decoding: browsers do not expose EXIF consistently.
  async function readExifOrientation(file) {
    if (!file || !/jpe?g/i.test(file.type || file.name || "")) return 1;
    const view = new DataView(await file.slice(0, 256 * 1024).arrayBuffer());
    if (view.byteLength < 4 || view.getUint16(0) !== 0xffd8) return 1;
    let offset = 2;
    while (offset + 4 < view.byteLength) {
      const marker = view.getUint16(offset); offset += 2;
      const length = view.getUint16(offset); offset += 2;
      if (marker === 0xffe1 && offset + length <= view.byteLength && view.getUint32(offset) === 0x45786966) {
        const tiff = offset + 6;
        const little = view.getUint16(tiff) === 0x4949;
        const first = tiff + view.getUint32(tiff + 4, little);
        const count = view.getUint16(first, little);
        for (let i = 0; i < count; i += 1) {
          const entry = first + 2 + i * 12;
          if (entry + 12 <= view.byteLength && view.getUint16(entry, little) === 0x0112) return view.getUint16(entry + 8, little) || 1;
        }
        return 1;
      }
      if (length < 2) break;
      offset += length - 2;
    }
    return 1;
  }

  async function orientedCanvas(file) {
    const orientation = await readExifOrientation(file);
    const image = await loadImage(await readAsDataUrl(file));
    const swap = orientation >= 5 && orientation <= 8;
    const canvas = document.createElement("canvas");
    canvas.width = swap ? image.naturalHeight : image.naturalWidth;
    canvas.height = swap ? image.naturalWidth : image.naturalHeight;
    const ctx = canvas.getContext("2d", { alpha: false, willReadFrequently: true });
    ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, canvas.width, canvas.height);
    const transforms = {
      2: [-1, 0, 0, 1, canvas.width, 0], 3: [-1, 0, 0, -1, canvas.width, canvas.height],
      4: [1, 0, 0, -1, 0, canvas.height], 5: [0, 1, 1, 0, 0, 0],
      6: [0, 1, -1, 0, canvas.width, 0], 7: [0, -1, -1, 0, canvas.width, canvas.height],
      8: [0, -1, 1, 0, 0, canvas.height]
    };
    if (transforms[orientation]) ctx.setTransform(...transforms[orientation]);
    ctx.drawImage(image, 0, 0);
    return canvas;
  }

  async function optimizeImage(file) {
    const image = await loadImage(await readAsDataUrl(file));
    const longestSide = Math.max(image.naturalWidth, image.naturalHeight);
    const scale = Math.min(1, 1600 / Math.max(1, longestSide));
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#fff"; context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    const dataUrl = canvas.toDataURL("image/jpeg", .86);
    return { dataUrl, fileName: String(file.name || "documento.jpg").replace(/\.[^.]+$/, "") + ".jpg", fileType: "image/jpeg", fileSize: Math.round(dataUrl.length * .75) };
  }

  function detectDocumentEdges(canvas) {
    const sample = document.createElement("canvas");
    const scale = Math.min(1, 700 / Math.max(canvas.width, canvas.height));
    sample.width = Math.round(canvas.width * scale); sample.height = Math.round(canvas.height * scale);
    const ctx = sample.getContext("2d", { willReadFrequently: true });
    ctx.drawImage(canvas, 0, 0, sample.width, sample.height);
    const pixels = ctx.getImageData(0, 0, sample.width, sample.height).data;
    const luminance = (x, y) => { const i = (y * sample.width + x) * 4; return .299 * pixels[i] + .587 * pixels[i + 1] + .114 * pixels[i + 2]; };
    const vertical = (x) => { let score = 0; for (let y = 2; y < sample.height - 2; y += 4) score += Math.abs(luminance(x, y) - luminance(x - 2, y)); return score / sample.height; };
    const horizontal = (y) => { let score = 0; for (let x = 2; x < sample.width - 2; x += 4) score += Math.abs(luminance(x, y) - luminance(x, y - 2)); return score / sample.width; };
    const best = (from, to, fn) => { let at = from; let value = -1; for (let p = from; p <= to; p += 2) { const s = fn(p); if (s > value) { value = s; at = p; } } return [at, value]; };
    const [left, ls] = best(Math.round(sample.width * .04), Math.round(sample.width * .42), vertical);
    const [right, rs] = best(Math.round(sample.width * .58), Math.round(sample.width * .96), vertical);
    const [top, ts] = best(Math.round(sample.height * .04), Math.round(sample.height * .42), horizontal);
    const [bottom, bs] = best(Math.round(sample.height * .58), Math.round(sample.height * .96), horizontal);
    const confidence = Math.min(1, (ls + rs + ts + bs) / 22);
    const sx = canvas.width / sample.width, sy = canvas.height / sample.height;
    return { confidence, corners: [{ x: left * sx, y: top * sy }, { x: right * sx, y: top * sy }, { x: right * sx, y: bottom * sy }, { x: left * sx, y: bottom * sy }] };
  }

  const distance = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
  function solveLinear(matrix, values) {
    for (let i = 0; i < values.length; i += 1) {
      let pivot = i; for (let r = i + 1; r < values.length; r += 1) if (Math.abs(matrix[r][i]) > Math.abs(matrix[pivot][i])) pivot = r;
      [matrix[i], matrix[pivot]] = [matrix[pivot], matrix[i]]; [values[i], values[pivot]] = [values[pivot], values[i]];
      const div = matrix[i][i] || 1e-10; for (let c = i; c < values.length; c += 1) matrix[i][c] /= div; values[i] /= div;
      for (let r = 0; r < values.length; r += 1) if (r !== i) { const m = matrix[r][i]; for (let c = i; c < values.length; c += 1) matrix[r][c] -= m * matrix[i][c]; values[r] -= m * values[i]; }
    }
    return values;
  }
  function homography(from, to) {
    const a = [], b = [];
    from.forEach((p, i) => { const q = to[i]; a.push([p.x,p.y,1,0,0,0,-q.x*p.x,-q.x*p.y]); b.push(q.x); a.push([0,0,0,p.x,p.y,1,-q.y*p.x,-q.y*p.y]); b.push(q.y); });
    return [...solveLinear(a, b), 1];
  }

  function perspectiveCorrect(canvas, selected) {
    const measuredW = Math.max(distance(selected[0], selected[1]), distance(selected[3], selected[2]));
    const measuredH = Math.max(distance(selected[0], selected[3]), distance(selected[1], selected[2]));
    // Pad to the physical ID-card ratio instead of cropping content from the detected quadrilateral.
    let contentW = Math.round(measuredW), contentH = Math.round(measuredH);
    const scale = Math.min(1, 3200 / Math.max(contentW, contentH)); contentW = Math.max(900, Math.round(contentW * scale)); contentH = Math.max(560, Math.round(contentH * scale));
    let outputW = contentW, outputH = contentH;
    if (outputW / outputH < IDENTITY_RATIO) outputW = Math.round(outputH * IDENTITY_RATIO); else outputH = Math.round(outputW / IDENTITY_RATIO);
    const offsetX = Math.floor((outputW - contentW) / 2), offsetY = Math.floor((outputH - contentH) / 2);
    const destination = [{x:offsetX,y:offsetY},{x:offsetX+contentW-1,y:offsetY},{x:offsetX+contentW-1,y:offsetY+contentH-1},{x:offsetX,y:offsetY+contentH-1}];
    const inverse = homography(destination, selected);
    const src = canvas.getContext("2d").getImageData(0, 0, canvas.width, canvas.height);
    const out = document.createElement("canvas"); out.width = outputW; out.height = outputH;
    const ctx = out.getContext("2d", { alpha: false }); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, outputW, outputH);
    const image = ctx.getImageData(0, 0, outputW, outputH), d = image.data, s = src.data;
    for (let y = offsetY; y < offsetY + contentH; y += 1) for (let x = offsetX; x < offsetX + contentW; x += 1) {
      const den = inverse[6]*x + inverse[7]*y + inverse[8]; const sx = (inverse[0]*x + inverse[1]*y + inverse[2])/den; const sy = (inverse[3]*x + inverse[4]*y + inverse[5])/den;
      const ix = Math.max(0, Math.min(canvas.width-1, Math.round(sx))), iy = Math.max(0, Math.min(canvas.height-1, Math.round(sy))); const si=(iy*canvas.width+ix)*4, di=(y*outputW+x)*4;
      for (let c=0;c<3;c+=1) d[di+c]=s[si+c]; d[di+3]=255;
    }
    // Moderate brightness/contrast plus an unsharp mask for legible small text.
    for (let i=0;i<d.length;i+=4) for (let c=0;c<3;c+=1) d[i+c]=Math.max(0,Math.min(255,(d[i+c]-128)*1.10+134));
    ctx.putImageData(image,0,0); ctx.filter="contrast(1.04) brightness(1.02)"; ctx.globalAlpha=.82; ctx.drawImage(out,0,0); ctx.globalAlpha=1; ctx.filter="none";
    return out;
  }

  function updatePreview() { const out = perspectiveCorrect(sourceCanvas, corners); cropPreview.src = out.toDataURL("image/jpeg", .92); return out; }
  function drawCropEditor() {
    const maxW = Math.min(window.innerWidth - 24, 920), maxH = Math.min(window.innerHeight * .55, 650), scale = Math.min(maxW/sourceCanvas.width,maxH/sourceCanvas.height,1);
    cropCanvas.width=Math.round(sourceCanvas.width*scale); cropCanvas.height=Math.round(sourceCanvas.height*scale); cropCanvas.getContext("2d").drawImage(sourceCanvas,0,0,cropCanvas.width,cropCanvas.height);
    cropHandles.innerHTML="";
    const points=corners.map((point,index)=>{ const handle=document.createElement("button"); handle.type="button"; handle.className="identity-crop-handle"; handle.setAttribute("aria-label",`Angolo ${index+1}`); handle.style.left=`${point.x*scale}px`; handle.style.top=`${point.y*scale}px`; handle.dataset.index=index; cropHandles.appendChild(handle); return `${point.x*scale},${point.y*scale}`; });
    cropHandles.style.setProperty("--crop-polygon", `polygon(${points.join(",")})`); cropHandles.dataset.scale=String(scale);
    updatePreview();
  }
  function closeCrop() { cropDialog.classList.add("hidden"); document.body.style.overflow=previousBodyOverflow; }
  async function openCrop(canvas) {
    stopCamera(); sourceCanvas=canvas; const detected=detectDocumentEdges(canvas); corners=detected.corners;
    byId("identity-crop-confidence").textContent=detected.confidence >= .55 ? "Bordi rilevati: controlla i quattro punti prima di continuare." : "Rilevamento incerto: posiziona manualmente i quattro punti sui bordi.";
    byId("identity-crop-confidence").classList.toggle("low-confidence",detected.confidence < .55);
    previousBodyOverflow=document.body.style.overflow; document.body.style.overflow="hidden"; cropDialog.classList.remove("hidden"); drawCropEditor();
  }

  cropHandles?.addEventListener("pointerdown", (event) => { const handle=event.target.closest(".identity-crop-handle"); if(!handle)return; activeHandle=Number(handle.dataset.index); handle.setPointerCapture(event.pointerId); });
  cropHandles?.addEventListener("pointermove", (event) => { if(activeHandle<0)return; const rect=cropCanvas.getBoundingClientRect(), scale=Number(cropHandles.dataset.scale)||1; corners[activeHandle]={x:Math.max(0,Math.min(sourceCanvas.width,(event.clientX-rect.left)/scale)),y:Math.max(0,Math.min(sourceCanvas.height,(event.clientY-rect.top)/scale))}; drawCropEditor(); });
  cropHandles?.addEventListener("pointerup",()=>{activeHandle=-1;});

  function stopCamera(){ if(stream){stream.getTracks().forEach((track)=>track.stop());stream=null;} if(video)video.srcObject=null; byId("identity-camera-snap").disabled=true; }
  async function startCamera(){ try { stopCamera(); stream=await navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:"environment"},width:{ideal:4096},height:{ideal:3072}},audio:false}); video.srcObject=stream; await video.play(); byId("identity-camera-snap").disabled=false; } catch(error){ showFeedback("Fotocamera non disponibile: usa il selettore del dispositivo.","error"); dedicatedInput.click(); } }
  function captureFrame(){ if(!video.videoWidth)return; const canvas=document.createElement("canvas"); canvas.width=video.videoWidth;canvas.height=video.videoHeight;canvas.getContext("2d",{alpha:false}).drawImage(video,0,0);openCrop(canvas); }
  function setIdentityMode(active){ identityPreset=active; flow?.classList.toggle("hidden",!active); genericFields.forEach((field)=>field?.classList.toggle("hidden",active)); form.classList.toggle("identity-preset-active",active); if(active){nameInput.value="Tessera di riconoscimento";noteInput.value="Documento personale di riconoscimento.";fileInput.value="";cameraInput.value="";processedIdentity=null;startCamera();}else stopCamera(); }

  tesseraButton?.addEventListener("click",()=>setIdentityMode(true));
  byId("private-docs-preset-pin-btn")?.addEventListener("click",()=>setIdentityMode(false));
  byId("identity-capture-cancel")?.addEventListener("click",()=>setIdentityMode(false));
  byId("identity-camera-start")?.addEventListener("click",startCamera);
  byId("identity-camera-snap")?.addEventListener("click",captureFrame);
  byId("identity-camera-fallback")?.addEventListener("click",()=>dedicatedInput.click());
  dedicatedInput?.addEventListener("change",async()=>{const file=dedicatedInput.files?.[0];if(file)openCrop(await orientedCanvas(file));});
  byId("identity-crop-close")?.addEventListener("click",closeCrop);
  byId("identity-crop-retry")?.addEventListener("click",()=>{closeCrop();processedIdentity=null;dedicatedInput.value="";startCamera();});
  byId("identity-crop-use")?.addEventListener("click",()=>{const out=updatePreview();const dataUrl=out.toDataURL("image/jpeg",.92);processedIdentity={dataUrl,fileName:"tessera-riconoscimento.jpg",fileType:"image/jpeg",fileSize:Math.round(dataUrl.length*.75)};closeCrop();showFeedback(`Foto pronta (${out.width} × ${out.height}px). Tocca l’anteprima per ingrandire.`,"success");});
  byId("identity-preview-zoom")?.addEventListener("click",()=>cropPreview.classList.toggle("is-zoomed"));
  window.addEventListener("orientationchange",()=>{if(sourceCanvas&&!cropDialog.classList.contains("hidden"))setTimeout(drawCropEditor,120);});

  async function prepareFile(file) {
    if (identityPreset) { if (!processedIdentity) throw new Error("Scatta, ritaglia e conferma la foto della tessera prima di salvare."); return processedIdentity; }
    if (!file) return { dataUrl:"",fileName:"",fileType:"",fileSize:0 };
    const type=String(file.type||"").toLowerCase(); if(type.startsWith("image/"))return optimizeImage(file);
    if(type!=="application/pdf")throw new Error("Puoi allegare soltanto immagini oppure PDF.");
    if(Number(file.size||0)>MAX_PDF_BYTES)throw new Error("Il PDF è troppo grande. Il limite è 500 KB.");
    return {dataUrl:await readAsDataUrl(file),fileName:file.name||"documento.pdf",fileType:type,fileSize:Number(file.size||0)};
  }

  form.addEventListener("submit",async(event)=>{
    event.preventDefault();event.stopImmediatePropagation();const user=firebase.auth().currentUser;if(!user)return showFeedback("Devi effettuare il login prima di salvare.","error");
    const name=String(nameInput.value||"").trim(),note=String(noteInput.value||"").trim(),file=fileInput.files?.[0]||cameraInput.files?.[0]||null;if(!name)return showFeedback("Inserisci la denominazione del documento.","error");
    saveButton.disabled=true;showFeedback("Preparazione e salvataggio del documento...");
    try{const prepared=await prepareFile(file);await firebase.firestore().collection("privateDocuments").doc(user.uid).collection("items").add({name,note,fileName:prepared.fileName,fileType:prepared.fileType,fileSize:prepared.fileSize,fileDataUrl:prepared.dataUrl,driveFileId:"",driveWebViewLink:"",storageMode:"private-firestore",ownerUid:user.uid,identityCardProcessed:identityPreset,createdAt:firebase.firestore.FieldValue.serverTimestamp(),updatedAt:firebase.firestore.FieldValue.serverTimestamp()});form.reset();setIdentityMode(false);showFeedback("Documento personale salvato correttamente.","success");}
    catch(error){console.error("Salvataggio documento personale non riuscito:",error);showFeedback(error?.message||"Salvataggio non riuscito. Riprova.","error");}finally{saveButton.disabled=false;}
  },true);
})();
