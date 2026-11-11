#!/usr/bin/env ruby
# Register the validated Firebase Apple configuration with the *generated*
# Capacitor Xcode project. Production web app and Android remain untouched.
require "xcodeproj"

root = File.expand_path("..", __dir__)
file_path = File.join(root, "ios", "App", "App", "GoogleService-Info.plist")
project_path = File.join(root, "ios", "App", "App.xcodeproj")
abort "No validated Firebase iOS plist" unless File.file?(file_path)
abort "No generated iOS Xcode project" unless File.directory?(project_path)

project = Xcodeproj::Project.open(project_path)
target = project.targets.find { |t| t.name == "App" && t.respond_to?(:resources_build_phase) }
abort "iOS App target not found" unless target

# main_group lives at ios/App, so 'App/...' is relative to it.
reference = project.main_group.files.find do |file|
  file.path == "App/GoogleService-Info.plist"
end
reference ||= project.main_group.new_file("App/GoogleService-Info.plist")

unless target.resources_build_phase.files_references.include?(reference)
  target.resources_build_phase.add_file_reference(reference)
end

project.save
puts "Firebase Apple plist registered in App bundle resources."
