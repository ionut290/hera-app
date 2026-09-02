(function installPwaWhazzupContinuousCamera() {
  "use strict";

  const isNative = Boolean(
    window.Capacitor &&
    typeof window.Capacitor.isNativePlatform === "function" &&
    window.Capacitor.isNativePlatform()
  );
  if (isNative || window.__heraPwaWhazzupContinuousCameraInstalled) return;
  window.__heraPwaWhazzupContinuousCameraInstalled = true;

  const supportsCamera = () => Boolean(
    window.isSecureContext &&
    navigator.mediaDevices &&
    typeof navigator.mediaDevices.getUserMedia === "function"
  );

  const waitForPhotoRuntime = () => new Promise((resolve) => {
    const startedAt = Date.now();
    const check = () => {
      if (
        typeof window.openWhazzupPhotoSourceChooser === "function" &&
        typeof window.pickWhazzupPhotos === "function" &&
        typeof window.saveWhazzupPhotoSelection === "function" &&
        typeof window.getWhazzupPhotos === "function" &&
        typeof window.getWhazzupPhotoNotes === "function"
      ) {
        resolve(true);
        return;
      }
      if (Date.now() - startedAt > 15000) {
        resolve(false);
        return;
      }
      window.setTimeout(check, 120);
    };
    check();
  });

  function makeOverlay() {
    const overlay = document.createElement("div");
    overlay.className = "hera-pwa-camera";
    overlay.setAttribute("role", "dialog");
    overlay.setAttribute("aria-modal", "true");
    overlay.innerHTML = `
      <style>
        .hera-pwa-camera{position:fixed;inset:0;z-index:2147483600;background:#000;color:#fff;display:flex;flex-direction:column;font-family:inherit;touch-action:none}
        .hera-pwa-camera__top{position:absolute;top:0;left:0;right:0;z-index:3;display:flex;align-items:center;justify-content:space-between;padding:calc(12px + env(safe-area-inset-top)) 16px 12px;background:linear-gradient(180deg,rgba(0,0,0,.72),transparent)}
        .hera-pwa-camera__close,.hera-pwa-camera__done,.hera-pwa-camera__switch{border:0;border-radius:999px;min-height:44px;padding:0 16px;background:rgba(28,28,30,.78);color:#fff;font:700 16px/1 inherit}
        .hera-pwa-camera__done{background:#13a860}
        .hera-pwa-camera__count{font-weight:800;font-size:17px;text-shadow:0 1px 3px #000}
        .hera-pwa-camera__stage{position:absolute;inset:0;overflow:hidden;background:#000}
        .hera-pwa-camera video{width:100%;height:100%;object-fit:cover;transform:scale(var(--hera-camera-digital-zoom,1));transform-origin:center;transition:transform .12s ease-out}
        .hera-pwa-camera__bottom{position:absolute;left:0;right:0;bottom:0;z-index:3;padding:14px 18px calc(18px + env(safe-area-inset-bottom));background:linear-gradient(0deg,rgba(0,0,0,.8),transparent);display:grid;grid-template-columns:72px 1fr 72px;align-items:end;gap:14px}
        .hera-pwa-camera__thumb{width:62px;height:62px;border-radius:12px;overflow:hidden;background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.35)}
        .hera-pwa-camera__thumb img{width:100%;height:100%;object-fit:cover}
        .hera-pwa-camera__shutter{justify-self:center;width:78px;height:78px;border-radius:50%;border:6px solid #fff;background:rgba(255,255,255,.25);box-shadow:0 0 0 2px rgba(0,0,0,.25) inset}
        .hera-pwa-camera__shutter:active{transform:scale(.94)}
        .hera-pwa-camera__switch{width:62px;padding:0;font-size:23px}
        .hera-pwa-camera__hint{position:absolute;left:50%;bottom:calc(164px + env(safe-area-inset-bottom));transform:translateX(-50%);z-index:3;background:rgba(0,0,0,.55);padding:7px 12px;border-radius:999px;font-size:13px;white-space:nowrap}
        .hera-pwa-camera__zoom{position:absolute;left:50%;bottom:calc(112px + env(safe-area-inset-bottom));transform:translateX(-50%);z-index:4;display:flex;align-items:center;gap:8px;background:rgba(0,0,0,.35);padding:5px;border-radius:999px}
        .hera-pwa-camera__zoom button{border:0;min-width:44px;height:36px;border-radius:999px;padding:0 10px;background:rgba(255,255,255,.16);color:#fff;font:800 14px/1 inherit}
        .hera-pwa-camera__zoom button.is-active{background:#fff;color:#111}
        .hera-pwa-camera__zoom-value{min-width:50px;text-align:center;font-weight:800;font-size:13px;text-shadow:0 1px 3px #000}
        .hera-pwa-camera__flash{position:absolute;inset:0;background:#fff;opacity:0;pointer-events:none;z-index:2}
        .hera-pwa-camera__flash.is-on{animation:heraPwaCameraFlash .18s ease-out}
        @keyframes heraPwaCameraFlash{0%{opacity:.7}100%{opacity:0}}
        @media (orientation:landscape){
          .hera-pwa-camera__top{padding:calc(8px + env(safe-area-inset-top)) calc(14px + env(safe-area-inset-right)) 8px calc(14px + env(safe-area-inset-left))}
          .hera-pwa-camera__bottom{top:0;right:0;bottom:0;left:auto;width:118px;padding:calc(72px + env(safe-area-inset-top)) calc(14px + env(safe-area-inset-right)) calc(16px + env(safe-area-inset-bottom)) 14px;background:linear-gradient(270deg,rgba(0,0,0,.78),transparent);display:flex;flex-direction:column;align-items:center;justify-content:flex-end;gap:16px}
          .hera-pwa-camera__shutter{order:1}
          .hera-pwa-camera__thumb{order:2;width:52px;height:52px}
          .hera-pwa-camera__switch{order:3;width:52px;min-height:44px}
          .hera-pwa-camera__zoom{left:auto;right:calc(126px + env(safe-area-inset-right));bottom:calc(18px + env(safe-area-inset-bottom));transform:none}
          .hera-pwa-camera__hint{left:50%;bottom:calc(18px + env(safe-area-inset-bottom));transform:translateX(-50%);font-size:12px}
        }
      </style>
      <div class="hera-pwa-camera__stage">
        <video playsinline autoplay muted></video>
        <div class="hera-pwa-camera__flash"></div>
      </div>
      <div class="hera-pwa-camera__top">
        <button class="hera-pwa-camera__close" type="button">Annulla</button>
        <span class="hera-pwa-camera__count">0 foto</span>
        <button class="hera-pwa-camera__done" type="button" disabled>Fine</button>
      </div>
      <div class="hera-pwa-camera__hint">Scatta tutte le foto, poi premi Fine</div>
      <div class="hera-pwa-camera__zoom" aria-label="Zoom fotocamera">
        <button type="button" data-camera-zoom="1" class="is-active">1×</button>
        <button type="button" data-camera-zoom="2">2×</button>
        <button type="button" data-camera-zoom="3">3×</button>
        <span class="hera-pwa-camera__zoom-value" aria-live="polite">1,0×</span>
      </div>
      <div class="hera-pwa-camera__bottom">
        <div class="hera-pwa-camera__thumb"></div>
        <button class="hera-pwa-camera__shutter" type="button" aria-label="Scatta foto"></button>
        <button class="hera-pwa-camera__switch" type="button" aria-label="Cambia fotocamera">↺</button>
      </div>`;
    return overlay;
  }

  async function startStream(video, facingMode) {
    const stream = await navigator.mediaDevices.getUserMedia({
      audio: false,
      video: {
        facingMode: { ideal: facingMode },
        width: { ideal: 1920 },
        height: { ideal: 1080 }
      }
    });
    video.srcObject = stream;
    await video.play();
    return stream;
  }

  function stopStream(stream) {
    try {
      stream?.getTracks?.().forEach((track) => track.stop());
    } catch (_) {}
  }

  function getZoomCapability(stream) {
    const track = stream?.getVideoTracks?.()[0];
    if (!track || typeof track.getCapabilities !== "function") return null;
    try {
      const caps = track.getCapabilities();
      const zoom = caps && caps.zoom;
      if (!zoom || !Number.isFinite(Number(zoom.min)) || !Number.isFinite(Number(zoom.max))) return null;
      return {
        track,
        min: Number(zoom.min),
        max: Number(zoom.max),
        step: Number(zoom.step || 0.1)
      };
    } catch (_) {
      return null;
    }
  }

  async function captureFrame(video, index, digitalZoom) {
    const sourceWidth = Number(video.videoWidth || 1280);
    const sourceHeight = Number(video.videoHeight || 720);
    const zoom = Math.max(1, Number(digitalZoom || 1));
    const cropWidth = Math.max(1, Math.round(sourceWidth / zoom));
    const cropHeight = Math.max(1, Math.round(sourceHeight / zoom));
    const sourceX = Math.max(0, Math.round((sourceWidth - cropWidth) / 2));
    const sourceY = Math.max(0, Math.round((sourceHeight - cropHeight) / 2));
    const maxSide = 1920;
    const scale = Math.min(1, maxSide / Math.max(cropWidth, cropHeight));
    const width = Math.max(1, Math.round(cropWidth * scale));
    const height = Math.max(1, Math.round(cropHeight * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    canvas.getContext("2d", { alpha: false }).drawImage(
      video,
      sourceX,
      sourceY,
      cropWidth,
      cropHeight,
      0,
      0,
      width,
      height
    );
    const blob = await new Promise((resolve, reject) => {
      canvas.toBlob(
        (value) => value ? resolve(value) : reject(new Error("Scatto non riuscito")),
        "image/jpeg",
        .92
      );
    });
    return new File([blob], `Foto-${String(index).padStart(2, "0")}.jpg`, {
      type: "image/jpeg",
      lastModified: Date.now()
    });
  }

  async function commitPhotos(impianto, button, files, options) {
    const current = window.getWhazzupPhotos(impianto).slice();
    const currentNotes = window.getWhazzupPhotoNotes(impianto).slice();
    let nextFiles = files.slice();
    let nextNotes = files.map(() => "");
    if (options.mode === "append") {
      nextFiles = [...current, ...files];
      nextNotes = [...currentNotes, ...files.map(() => "")];
    } else if (options.mode === "replace-one") {
      nextFiles = current.slice();
      nextNotes = currentNotes.slice();
      if (files[0]) nextFiles.splice(Number(options.index || 0), 1, files[0]);
    }
    await window.saveWhazzupPhotoSelection(impianto, nextFiles, nextNotes);
    if (typeof window.updateWhazzupAttachmentButton === "function") {
      window.updateWhazzupAttachmentButton(button, impianto);
    }
    if (typeof window.openWhazzupPhotoManager === "function") {
      window.openWhazzupPhotoManager(impianto, button);
    }
  }

  async function openContinuousCamera(impianto, button, options = {}) {
    if (!supportsCamera()) {
      window.pickWhazzupPhotos(impianto, button, { ...options, source: "camera" });
      return;
    }

    const overlay = makeOverlay();
    const video = overlay.querySelector("video");
    const shutter = overlay.querySelector(".hera-pwa-camera__shutter");
    const done = overlay.querySelector(".hera-pwa-camera__done");
    const close = overlay.querySelector(".hera-pwa-camera__close");
    const count = overlay.querySelector(".hera-pwa-camera__count");
    const thumb = overlay.querySelector(".hera-pwa-camera__thumb");
    const switchButton = overlay.querySelector(".hera-pwa-camera__switch");
    const flash = overlay.querySelector(".hera-pwa-camera__flash");
    const zoomBar = overlay.querySelector(".hera-pwa-camera__zoom");
    const zoomValue = overlay.querySelector(".hera-pwa-camera__zoom-value");
    const zoomButtons = [...overlay.querySelectorAll("[data-camera-zoom]")];
    const files = [];
    const urls = [];
    let stream = null;
    let facingMode = "environment";
    let busy = false;
    let closed = false;
    let desiredZoom = 1;
    let digitalZoom = 1;
    let pinchStartDistance = 0;
    let pinchStartZoom = 1;

    const cleanup = () => {
      if (closed) return;
      closed = true;
      stopStream(stream);
      urls.forEach((url) => URL.revokeObjectURL(url));
      overlay.remove();
    };

    const refresh = () => {
      const n = files.length;
      count.textContent = `${n} foto`;
      done.disabled = n === 0;
      const last = urls[urls.length - 1];
      thumb.innerHTML = last ? `<img src="${last}" alt="Ultima foto">` : "";
    };

    const applyZoom = async (requestedZoom) => {
      desiredZoom = Math.max(1, Math.min(3, Number(requestedZoom || 1)));
      const capability = getZoomCapability(stream);
      let nativeZoomApplied = false;
      if (capability && capability.max > capability.min) {
        const nativeZoom = Math.max(capability.min, Math.min(capability.max, desiredZoom));
        try {
          await capability.track.applyConstraints({ advanced: [{ zoom: nativeZoom }] });
          nativeZoomApplied = Math.abs(nativeZoom - desiredZoom) < 0.08;
        } catch (_) {}
      }
      digitalZoom = nativeZoomApplied ? 1 : desiredZoom;
      overlay.style.setProperty("--hera-camera-digital-zoom", String(digitalZoom));
      zoomValue.textContent = `${desiredZoom.toFixed(1).replace(".", ",")}×`;
      zoomButtons.forEach((zoomButton) => {
        const value = Number(zoomButton.dataset.cameraZoom || 1);
        zoomButton.classList.toggle("is-active", Math.abs(value - desiredZoom) < 0.15);
      });
    };

    const touchDistance = (touches) => {
      if (!touches || touches.length < 2) return 0;
      const dx = touches[0].clientX - touches[1].clientX;
      const dy = touches[0].clientY - touches[1].clientY;
      return Math.hypot(dx, dy);
    };

    document.body.appendChild(overlay);
    try {
      stream = await startStream(video, facingMode);
      await applyZoom(1);
    } catch (error) {
      cleanup();
      console.warn("Fotocamera continua PWA non disponibile, uso fallback:", error);
      window.pickWhazzupPhotos(impianto, button, { ...options, source: "camera" });
      return;
    }

    close.addEventListener("click", cleanup);

    zoomBar.addEventListener("click", (event) => {
      const zoomButton = event.target.closest("[data-camera-zoom]");
      if (!zoomButton || busy) return;
      applyZoom(Number(zoomButton.dataset.cameraZoom || 1));
    });

    overlay.addEventListener("touchstart", (event) => {
      if (event.touches.length !== 2) return;
      pinchStartDistance = touchDistance(event.touches);
      pinchStartZoom = desiredZoom;
    }, { passive: true });

    overlay.addEventListener("touchmove", (event) => {
      if (event.touches.length !== 2 || !pinchStartDistance) return;
      const distance = touchDistance(event.touches);
      const nextZoom = pinchStartZoom * (distance / pinchStartDistance);
      applyZoom(nextZoom);
    }, { passive: true });

    overlay.addEventListener("touchend", (event) => {
      if (event.touches.length < 2) pinchStartDistance = 0;
    }, { passive: true });

    shutter.addEventListener("click", async () => {
      if (busy) return;
      busy = true;
      shutter.disabled = true;
      try {
        const file = await captureFrame(video, files.length + 1, digitalZoom);
        if (options.mode === "replace-one") {
          files.splice(0, files.length, file);
          urls.splice(0).forEach((url) => URL.revokeObjectURL(url));
        } else {
          files.push(file);
        }
        const url = URL.createObjectURL(file);
        urls.push(url);
        flash.classList.remove("is-on");
        void flash.offsetWidth;
        flash.classList.add("is-on");
        refresh();
      } catch (error) {
        window.alert(error?.message || "Non riesco a scattare la foto.");
      } finally {
        busy = false;
        shutter.disabled = false;
      }
    });

    done.addEventListener("click", async () => {
      if (!files.length || busy) return;
      busy = true;
      done.disabled = true;
      try {
        const selected = files.slice();
        cleanup();
        await commitPhotos(impianto, button, selected, options);
      } catch (error) {
        console.error("Salvataggio foto PWA non riuscito:", error);
        window.alert(error?.message || "Non riesco a salvare le foto.");
      }
    });

    switchButton.addEventListener("click", async () => {
      if (busy) return;
      busy = true;
      switchButton.disabled = true;
      const nextFacing = facingMode === "environment" ? "user" : "environment";
      try {
        stopStream(stream);
        stream = await startStream(video, nextFacing);
        facingMode = nextFacing;
        await applyZoom(1);
      } catch (error) {
        try {
          stream = await startStream(video, facingMode);
          await applyZoom(1);
        } catch (_) {}
        console.warn("Cambio fotocamera PWA non disponibile:", error);
      } finally {
        busy = false;
        switchButton.disabled = false;
      }
    });
  }

  async function installChooserOverride() {
    if (!await waitForPhotoRuntime()) return;
    if (window.__heraOriginalWhazzupPhotoSourceChooser) return;
    window.__heraOriginalWhazzupPhotoSourceChooser = window.openWhazzupPhotoSourceChooser;
    window.openWhazzupPhotoSourceChooser = function (impianto, button, options = {}) {
      const overlay = document.createElement("div");
      overlay.className = "whazzup-photo-manager whazzup-photo-source-chooser";
      overlay.setAttribute("role", "dialog");
      overlay.setAttribute("aria-modal", "true");
      overlay.innerHTML = `
        <section class="whazzup-photo-source-card">
          <header>
            <p class="management-eyebrow">AGGIUNGI FOTO</p>
            <h2>Come vuoi aggiungere le foto?</h2>
            <p>Con la fotocamera puoi fare più scatti consecutivi senza uscire.</p>
          </header>
          <div class="whazzup-photo-source-actions">
            <button type="button" class="btn whazzup-photo-camera-btn" data-pwa-photo-source="camera">
              <span aria-hidden="true">📷</span><strong>Scatta foto</strong><small>Fotocamera continua</small>
            </button>
            <button type="button" class="btn whazzup-photo-gallery-btn" data-pwa-photo-source="gallery">
              <span aria-hidden="true">🖼️</span><strong>Scegli dalla galleria</strong><small>Una o più foto salvate</small>
            </button>
          </div>
          <button type="button" class="btn whazzup-photo-source-cancel" data-pwa-photo-source="cancel">Annulla</button>
        </section>`;
      const dismiss = () => overlay.remove();
      overlay.addEventListener("click", (event) => {
        if (event.target === overlay) return dismiss();
        const action = event.target.closest("[data-pwa-photo-source]")?.dataset.pwaPhotoSource;
        if (!action) return;
        dismiss();
        if (action === "camera") openContinuousCamera(impianto, button, options);
        else if (action === "gallery") {
          window.pickWhazzupPhotos(impianto, button, { ...options, source: "gallery" });
        }
      });
      document.body.appendChild(overlay);
      overlay.querySelector("[data-pwa-photo-source='camera']")?.focus();
    };
  }

  installChooserOverride().catch((error) => {
    console.warn("Fotocamera continua PWA non installata:", error);
  });
})();
