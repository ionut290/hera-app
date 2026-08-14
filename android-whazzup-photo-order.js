(function installAndroidWhazzupPhotoOrderFix() {
  "use strict";

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
    const dedicatedPlugin = getDedicatedAndroidWhazzupPhotoPlugin();
    if (dedicatedPlugin) {
      await sharePhotosThroughDedicatedPlugin(dedicatedPlugin, orderedFiles);
      if (!safeOpenWhatsAppMessage(message)) {
        throw new Error("Impossibile aprire il messaggio Whazzup finale");
      }
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
        dialogTitle: "Condividi prima tutte le foto su Whazzup"
      });
      scheduleNativeWhazzupShareCleanup(plugins.filesystem, folderPath);
      if (!safeOpenWhatsAppMessage(message)) {
        throw new Error("Impossibile aprire il messaggio Whazzup finale");
      }
      return true;
    } catch (error) {
      await removeNativeWhazzupShareFolder(plugins.filesystem, folderPath);
      throw error;
    }
  }

  window.shareWhazzupPhotosDedicatedAndroid = sharePhotosThroughDedicatedPlugin;
  window.shareWhazzupPhotosNativeAndroid = shareWhazzupPhotosNativeAndroidInOrder;
})();
