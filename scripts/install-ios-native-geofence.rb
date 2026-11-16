#!/usr/bin/env ruby
# Install iOS-only Core Location background geofencing into the generated Capacitor project.
# Run after "npx cap add ios" and "npx cap sync ios", before xcodebuild.
require "fileutils"
require "xcodeproj"
require "rexml/document"

root = File.expand_path("..", __dir__)
app_dir = File.join(root, "ios", "App", "App")
project_path = File.join(root, "ios", "App", "App.xcodeproj")
native_dir = File.join(root, "ios-native")
files = %w[VargaGeofenceManager.swift VargaGeofencePlugin.swift VargaBridgeViewController.swift]
abort "Generated iOS App target missing" unless Dir.exist?(app_dir) && Dir.exist?(project_path)

files.each do |name|
  src = File.join(native_dir, name)
  abort "Native geofence source missing: #{name}" unless File.file?(src)
  FileUtils.cp(src, File.join(app_dir, name))
end

app_delegate_path = File.join(app_dir, "AppDelegate.swift")
app_delegate = File.read(app_delegate_path)
boot = "VargaGeofenceManager.shared.bootstrap()"
unless app_delegate.include?(boot)
  anchor = "// Override point for customization after application launch."
  abort "Capacitor AppDelegate launch hook changed: refusing unsafe patch" unless app_delegate.scan(anchor).length == 1
  app_delegate = app_delegate.sub(anchor, "#{boot}\n        #{anchor}")
  File.write(app_delegate_path, app_delegate)
end

storyboard_path = File.join(app_dir, "Base.lproj", "Main.storyboard")
abort "Capacitor storyboard missing" unless File.file?(storyboard_path)
storyboard = File.read(storyboard_path)
unless storyboard.include?('customClass="VargaBridgeViewController"')
  old = 'customClass="CAPBridgeViewController"'
  abort "Unexpected storyboard root controller" unless storyboard.scan(old).length == 1
  storyboard = storyboard.sub(old, 'customClass="VargaBridgeViewController"')
  # The class is compiled into the app, not the Capacitor framework.
  storyboard = storyboard.sub(/(customClass="VargaBridgeViewController"[^>]*?)customModule="Capacitor"/, '\\1customModule="App"')
  File.write(storyboard_path, storyboard)
end
REXML::Document.new(File.read(storyboard_path))

project = Xcodeproj::Project.open(project_path)
target = project.targets.find { |t| t.name == "App" && t.respond_to?(:source_build_phase) }
abort "App source target not found" unless target
files.each do |name|
  rel = "App/#{name}"
  reference = project.main_group.files.find { |f| f.path == rel }
  reference ||= project.main_group.new_file(rel)
  target.source_build_phase.add_file_reference(reference) unless target.source_build_phase.files_references.include?(reference)
end
project.save

files.each do |name|
  abort "Geofencing source not registered" unless target.source_build_phase.files_references.any? { |f| f.path == "App/#{name}" }
end
abort "Cold-start hook missing" unless File.read(app_delegate_path).include?(boot)
abort "Custom Capacitor bridge missing" unless File.read(storyboard_path).include?('customClass="VargaBridgeViewController"')
puts "Geofencing installed: native Swift files + cold-start AppDelegate + custom Capacitor bridge."
