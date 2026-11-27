import json
import plistlib
import tempfile
import unittest
from pathlib import Path
from prepare_v2_native_social import prepare,validate_config,PROJECT,BUNDLE,APP_ID,SENDER

def sample():
    client=SENDER+"-fixture.apps.googleusercontent.com"
    return {"PROJECT_ID":PROJECT,"BUNDLE_ID":BUNDLE,"GOOGLE_APP_ID":APP_ID,"GCM_SENDER_ID":SENDER,"CLIENT_ID":client,"REVERSED_CLIENT_ID":".".join(reversed(client.split("."))),"API_KEY":"AIza"+"a"*35}

class NativeConfigTests(unittest.TestCase):
    def test_exact_identity(self):
        self.assertEqual(validate_config(sample()),sample()["REVERSED_CLIENT_ID"])
        for field in ("PROJECT_ID","BUNDLE_ID","GOOGLE_APP_ID","GCM_SENDER_ID","CLIENT_ID","REVERSED_CLIENT_ID","API_KEY"):
            data=sample();data[field]="unexpected"
            with self.subTest(field=field),self.assertRaises(ValueError):validate_config(data)
    def test_config_preserves_other_url_types_and_is_idempotent(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);app=root/"ios/App/App";app.mkdir(parents=True)
            config=root/"capacitor.config.json";config.write_text(json.dumps({"appId":BUNDLE,"appName":"V2","webDir":"www"}))
            plist=app/"Info.plist";plist.write_bytes(plistlib.dumps({"CFBundleURLTypes":[{"CFBundleURLSchemes":["existing"]}]}))
            pod=root/"ios/App/Podfile";pod.write_text("target 'App' do\n  capacitor_pods\n  # Add your Pods here\nend\npost_install do |installer|\n  assertDeploymentTarget(installer)\nend\n")
            src=root/"client.plist";src.write_bytes(plistlib.dumps(sample()))
            prepare(root,src);prepare(root,src)
            self.assertEqual(len(plistlib.loads(plist.read_bytes())["CFBundleURLTypes"]),2)
            self.assertEqual(plistlib.loads((app/"GoogleService-Info.plist").read_bytes()),sample())
            self.assertEqual(plistlib.loads((app/"App.entitlements").read_bytes())["com.apple.developer.applesignin"],["Default"])
            self.assertEqual(json.loads(config.read_text())["plugins"]["FirebaseAuthentication"]["providers"],["google.com","apple.com"])
            self.assertEqual(pod.read_text().count("CapacitorFirebaseAuthentication/Google"),1)
            self.assertEqual(pod.read_text().count("CODE_SIGNING_ALLOWED"),1)
    def test_wrong_project_is_untouched(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);src=root/"client.plist";data=sample();data["PROJECT_ID"]="other";src.write_bytes(plistlib.dumps(data))
            with self.assertRaises(ValueError):prepare(root,src)
            self.assertFalse((root/"ios").exists())

if __name__=="__main__":unittest.main()
