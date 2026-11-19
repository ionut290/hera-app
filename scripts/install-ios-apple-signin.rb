#!/usr/bin/env ruby
# Configure Sign in with Apple in the TEMPORARY generated Capacitor iOS target.
# Run after 'npx cap add ios && npx cap sync ios', before signing or archiving.
# Never change Pods or Apple certificates.
require "fileutils"
require "xcodeproj"
require "rexml/document"

root = File.expand_path("..", __dir__)
project_path = File.join(root, "ios", "App", "App.xcodeproj")
app_dir = File.join(root, "ios", "App", "App")
entitlements_path = File.join(app_dir, "App.entitlements")
abort "Generated Capacitor iOS project missing" unless Dir.exist?(project_path) && Dir.exist?(app_dir)

xml = <<~XML
  <?xml version="1.0" encoding="UTF-8"?>
  <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
  <plist version="1.0">
    <dict>
      <key>com.apple.developer.applesignin</key>
      <array><string>Default</string></array>
    </dict>
  </plist>
XML

# Avoid destroying any entitlements that other capabilities may have added.
if File.exist?(entitlements_path)
  existing = File.read(entitlements_path)
  abort "Existing App.entitlements contains unrelated capabilities; merge manually" unless
    existing.include?("com.apple.developer.applesignin") &&
    existing.include?("<string>Default</string>")
else
  File.write(entitlements_path, xml)
end
REXML::Document.new(File.read(entitlements_path))

project = Xcodeproj::Project.open(project_path)
matches = project.targets.select { |t| t.name == "App" && t.respond_to?(:build_configurations) }
abort "Exactly one App target required" unless matches.size == 1
app = matches.first

app.build_configurations.each do |configuration|
  previous = configuration.build_settings["CODE_SIGN_ENTITLEMENTS"]
  if previous && previous != "App/App.entitlements"
    abort "Existing signing entitlements at #{previous}; refusing to overwrite"
  end
  configuration.build_settings["CODE_SIGN_ENTITLEMENTS"] = "App/App.entitlements"
end

attributes = project.root_object.attributes["TargetAttributes"] ||= {}
attributes[app.uuid] ||= {}
capabilities = attributes[app.uuid]["SystemCapabilities"] ||= {}
capabilities["com.apple.SignInWithApple"] = { "enabled" => 1 }
project.save

check = Xcodeproj::Project.open(project_path)
target = check.targets.find { |t| t.name == "App" }
abort "Missing Sign in with Apple build setting" unless target.build_configurations.all? do |c|
  c.build_settings["CODE_SIGN_ENTITLEMENTS"] == "App/App.entitlements"
end
abort "Sign in with Apple entitlement missing" unless File.read(entitlements_path)
  .include?("com.apple.developer.applesignin")
puts "Sign in with Apple installed in generated iOS App target (Debug and Release)."
