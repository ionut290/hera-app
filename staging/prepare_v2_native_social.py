#!/usr/bin/env python3
"""Configure only the independent V2 native project; never print client config."""
import json
import plistlib
import re
import sys
from pathlib import Path

PROJECT = "varga-cantieri-v2-staging"
BUNDLE = "it.vargacantieri.staging"
APP_ID = "1:394791335182:ios:b3634467e6741c1981935f"
SENDER = "394791335182"

def validate_config(config):
    for key, expected in {"PROJECT_ID":PROJECT,"BUNDLE_ID":BUNDLE,"GOOGLE_APP_ID":APP_ID,"GCM_SENDER_ID":SENDER}.items():
        if config.get(key) != expected:
            raise ValueError("Unexpected Firebase iOS identity: " + key)
    client = config.get("CLIENT_ID", "")
    if not isinstance(client,str) or not re.fullmatch(SENDER+r"-[A-Za-z0-9_-]+\.apps\.googleusercontent\.com",client):
        raise ValueError("Missing staging Google OAuth iOS client")
    reverse = ".".join(reversed(client.split(".")))
    if config.get("REVERSED_CLIENT_ID") != reverse:
        raise ValueError("Google callback URL scheme mismatch")
    if not re.fullmatch(r"AIza[A-Za-z0-9_-]{35}",config.get("API_KEY","")):
        raise ValueError("Missing Firebase iOS API key")
    return reverse

def prepare(root, source):
    root,source=Path(root),Path(source)
    raw=source.read_bytes()
    config=plistlib.loads(raw)
    scheme=validate_config(config)
    cap_path=root/"capacitor.config.json"
    cap=json.loads(cap_path.read_text())
    if cap.get("appId")!=BUNDLE: raise ValueError("Wrong Capacitor project")
    app=root/"ios/App/App"
    info=plistlib.loads((app/"Info.plist").read_bytes())
    # No unrelated URL types are removed.
    urls=info.setdefault("CFBundleURLTypes",[])
    if not any(scheme in x.get("CFBundleURLSchemes",[]) for x in urls):
        urls.append({"CFBundleTypeRole":"Editor","CFBundleURLSchemes":[scheme]})
    cap.setdefault("plugins",{})["FirebaseAuthentication"]={"skipNativeAuth":True,"providers":["google.com","apple.com"]}
    pod=root/"ios/App/Podfile"
    podtext=pod.read_text()
    target="# Add your Pods here"
    post="post_install do |installer|"
    if podtext.count(target)!=1 or podtext.count(post)!=1: raise ValueError("Unexpected Capacitor Podfile")
    google="  pod 'CapacitorFirebaseAuthentication/Google', :path => '../../node_modules/@capacitor-firebase/authentication'"
    if google not in podtext: podtext=podtext.replace(target,target+"\n"+google)
    bundle_signing="""  installer.pods_project.targets.each do |target|
    if target.respond_to?(:product_type) && target.product_type == 'com.apple.product-type.bundle'
      target.build_configurations.each { |config| config.build_settings['CODE_SIGNING_ALLOWED'] = 'NO' }
    end
  end"""
    if bundle_signing not in podtext: podtext=podtext.replace(post,post+"\n"+bundle_signing)
    # A generated, app-specific target entitlement, not an entitlement bypass.
    ent=app/"App.entitlements"
    existing=plistlib.loads(ent.read_bytes()) if ent.exists() else {}
    existing["com.apple.developer.applesignin"]=["Default"]
    (app/"Info.plist").write_bytes(plistlib.dumps(info))
    (app/"GoogleService-Info.plist").write_bytes(raw)
    ent.write_bytes(plistlib.dumps(existing))
    cap_path.write_text(json.dumps(cap,indent=2)+"\n")
    pod.write_text(podtext)
    return True

if __name__=="__main__":
    try:
        prepare(sys.argv[1],sys.argv[2])
        print("Validated and configured exact V2 iOS social provider files.")
    except (ValueError,TypeError,KeyError,OSError,IndexError,plistlib.InvalidFileException):
        sys.exit("V2 native social configuration failed; no signing permitted.")
