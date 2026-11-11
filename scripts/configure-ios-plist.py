#!/usr/bin/env python3
"""Configure purpose strings on a newly generated Capacitor iOS app.

Only the generated ios/App/App/Info.plist is changed. The production web app
and Android project are never edited by this helper.
"""

from pathlib import Path
import plistlib

ROOT = Path(__file__).resolve().parent.parent
INFO_PLIST = ROOT / "ios" / "App" / "App" / "Info.plist"

PURPOSE_STRINGS = {
    "NSCameraUsageDescription": (
        "Varga Cantieri utilizza la fotocamera per fotografare impianti, "
        "interventi e documenti quando lo richiedi."
    ),
    "NSPhotoLibraryUsageDescription": (
        "Varga Cantieri accede alle foto selezionate per allegarle a "
        "cantieri, interventi e documenti."
    ),
    "NSPhotoLibraryAddUsageDescription": (
        "Varga Cantieri salva nella libreria foto soltanto i contenuti "
        "che scegli di esportare."
    ),
    "NSLocationWhenInUseUsageDescription": (
        "Varga Cantieri utilizza la posizione mentre usi l'app per "
        "mostrare i cantieri vicini e le indicazioni sulla mappa."
    ),
    "NSMicrophoneUsageDescription": (
        "Varga Cantieri utilizza il microfono soltanto quando scegli "
        "di registrare contenuti audio."
    ),
}


def main() -> None:
    if not INFO_PLIST.is_file():
        raise SystemExit(
            "Progetto iOS mancante: prima eseguire 'npx cap add ios' "
            "e 'npx cap sync ios'."
        )

    with INFO_PLIST.open("rb") as file:
        data = plistlib.load(file)

    if not isinstance(data, dict):
        raise SystemExit("Info.plist non valido: formato inatteso.")

    data["CFBundleDisplayName"] = "VARGA CANTIERI"
    data.update(PURPOSE_STRINGS)

    with INFO_PLIST.open("wb") as file:
        plistlib.dump(data, file, sort_keys=False)

    with INFO_PLIST.open("rb") as file:
        verified = plistlib.load(file)
    for key, value in PURPOSE_STRINGS.items():
        if verified.get(key) != value:
            raise SystemExit(f"Permesso iOS non configurato: {key}")

    print("Info.plist iOS predisposto (permessi solo con app in uso).")


if __name__ == "__main__":
    main()
