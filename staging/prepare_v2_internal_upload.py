#!/usr/bin/env python3
"""Explicitly gated, internal TestFlight upload options for the separate V2 app."""
import os
import plistlib
import sys
from pathlib import Path
from prepare_v2_export import options as local_options

def options(authorized):
    if authorized != "true":
        raise ValueError("Internal V2 upload has not been enabled")
    result = local_options()
    result["destination"] = "upload"
    return result

if __name__ == "__main__":
    root = Path(os.environ["RUNNER_TEMP"]).resolve()
    destination = Path(sys.argv[1]).resolve()
    if destination.parent != root or destination.name != "VargaV2InternalUploadOptions.plist":
        sys.exit("Only the expected private temporary options path is allowed.")
    with destination.open("xb") as file:
        plistlib.dump(options(os.environ.get("V2_INTERNAL_UPLOAD_READY")), file)
    print("Explicitly enabled internal-only upload options for V2.")
