(function installAndroidWhazzupPhotoOrderFix() {
  "use strict";

  const PHOTO_SHARE_AFTER_MESSAGE_DELAY_MS = 8000;

  function waitBeforeWhazzupPhotos() {
    return new Promise((resolve) => {
      window.setTimeout(resolve, PHOTO_SHARE_AFTER_MESSAGE_DELAY_MS);
    });
  }

  async function sharePhotosThroughDedicatedPlugin(plugin, orderedFiles) {
    let sessionId = "";
    try {
      const session = await plugin.begin();
      sessionId = String(session?.sessionId || "");
      if (!sessionId) throw new Error("Sessione foto Android non disponibile");

      for (const file of orderedFiles) {
        await plugin.addPhoto({
          sessionId,
          fileName: String(file?.name || "foto.jpg"),
          mimeType: String(file?.type || "image/jpeg"),
          data: await readWhazzupPhotoAsBase64(file)
        });
      }

      return await plugin.share({ sessionId });
    } catch (error) {
      if (sessionId && typeof plugin.discard === "function") {
        try {
          await plugin.discard({ sessionId });
        } catch (_) {}
      }
      throw error;
    }
  }

  async function shareWhazzupPhotosNativeAndroidInOrder(orderedFiles, message) {
    if (!safeOpenWhatsAppMessage(message)) {
      throw new Error("Impossibile aprire il messaggio Whazzup iniziale");
    }

    await waitBeforeWhazzupPhotos();

    const dedicatedPlugin = getDedicatedAndroidWhazzupPhotoPlugin();
    if (dedicatedPlugin) {
      await sharePhotosThroughDedicatedPlugin(dedicatedPlugin, orderedFiles);
      return true;
    }

    const plugins = getNativeAndroidWhazzupSharePlugins();
    if (!plugins) return null;
    const folderPath = `hera-whazzup-share/${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    const fileUris = [];
    try {
      for (const file of orderedFiles) {
        const fileName = String(file?.name || `Foto-${String(fileUris.length + 1).padStart(2, "0")}.jpg`);
        const result = await plugins.filesystem.writeFile({
          path: `${folderPath}/${fileName}`,
          data: await readWhazzupPhotoAsBase64(file),
          directory: "CACHE",
          recursive: true
        });
        fileUris.push(result.uri);
      }
      await plugins.share.share({
        files: fileUris,
        dialogTitle: "Condividi le foto su Whazzup"
      });
      scheduleNativeWhazzupShareCleanup(plugins.filesystem, folderPath);
      return true;
    } catch (error) {
      await removeNativeWhazzupShareFolder(plugins.filesystem, folderPath);
      throw error;
    }
  }

  window.shareWhazzupPhotosDedicatedAndroid = sharePhotosThroughDedicatedPlugin;
  window.shareWhazzupPhotosNativeAndroid = shareWhazzupPhotosNativeAndroidInOrder;
})();

(function installWhazzupContinuousCamera() {
  "use strict";

  const isNativeAndroid = Boolean(
    window.Capacitor?.isNativePlatform?.()
    && window.Capacitor?.getPlatform?.() === "android"
  );
  if (!isNativeAndroid || window.__heraWhazzupContinuousCameraInstalled) return;

  const cameraPlugin = window.Capacitor?.Plugins?.HeraContinuousCamera
    || window.Capacitor?.registerPlugin?.("HeraContinuousCamera")
    || null;
  if (!cameraPlugin) return;

  const filePathToImageFile = async (photo, index) => {
    const path = String(photo?.path || "").trim();
    if (!path) return null;
    const webPath = window.Capacitor.convertFileSrc(path);
    const response = await fetch(webPath, { cache: "no-store" });
    if (!response.ok) throw new Error(`Foto ${index + 1} non leggibile`);
    const blob = await response.blob();
    const fileName = String(photo?.name || `Foto-${String(index + 1).padStart(2, "0")}.jpg`);
    return new File([blob], fileName, {
      type: String(photo?.type || blob.type || "image/jpeg"),
      lastModified: Date.now()
    });
  };

  const install = () => {
    if (window.__heraWhazzupContinuousCameraInstalled) return true;
    if (typeof window.pickWhazzupPhotos !== "function") return false;
    if (typeof window.getWhazzupPhotos !== "function") return false;
    if (typeof window.getWhazzupPhotoNotes !== "function") return false;
    if (typeof window.saveWhazzupPhotoSelection !== "function") return false;

    const legacyPickWhazzupPhotos = window.pickWhazzupPhotos;

    window.pickWhazzupPhotos = async function pickWhazzupPhotosContinuous(impianto, button, options = {}) {
      if (options.source !== "camera") {
        return legacyPickWhazzupPhotos(impianto, button, options);
      }

      try {
        const maxPhotos = options.mode === "replace-one" ? 1 : 10;
        const result = await cameraPlugin.capture({ maxPhotos });
        if (result?.cancelled) return;

        const nativePhotos = Array.isArray(result?.photos) ? result.photos : [];
        if (!nativePhotos.length) return;

        const capturedFiles = (await Promise.all(
          nativePhotos.map((photo, index) => filePathToImageFile(photo, index))
        )).filter(Boolean);
        if (!capturedFiles.length) return;

        const currentFiles = window.getWhazzupPhotos(impianto).slice();
        const currentNotes = window.getWhazzupPhotoNotes(impianto).slice();
        let nextFiles = capturedFiles;
        let nextNotes = capturedFiles.map(() => "");

        if (options.mode === "append") {
          nextFiles = [...currentFiles, ...capturedFiles];
          nextNotes = [...currentNotes, ...capturedFiles.map(() => "")];
        } else if (options.mode === "replace-one") {
          const index = Math.max(0, Number(options.index || 0));
          nextFiles = currentFiles.slice();
          nextFiles.splice(index, 1, capturedFiles[0]);
          nextNotes = currentNotes;
        }

        await window.saveWhazzupPhotoSelection(impianto, nextFiles, nextNotes);
        if (typeof window.updateWhazzupAttachmentButton === "function") {
          window.updateWhazzupAttachmentButton(button, impianto);
        }
        if (options.reopenManager !== false && typeof window.openWhazzupPhotoManager === "function") {
          window.openWhazzupPhotoManager(impianto, button);
        }
      } catch (error) {
        console.warn("Fotocamera continua Whazzup non disponibile; uso il flusso standard.", error);
        if (typeof window.showToast === "function") {
          window.showToast("Fotocamera continua non disponibile: apro quella standard.");
        }
        return legacyPickWhazzupPhotos(impianto, button, options);
      }
    };

    window.__heraWhazzupContinuousCameraInstalled = true;
    return true;
  };

  if (install()) return;
  let attempts = 0;
  const timer = window.setInterval(() => {
    attempts += 1;
    if (install() || attempts >= 40) window.clearInterval(timer);
  }, 250);
})();
