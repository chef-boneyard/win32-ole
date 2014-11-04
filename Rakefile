require 'rake'
require 'rake/testtask'
require 'rake/clean'

CLEAN.include("*.gem", "*.rbc")

Rake::TestTask.new do |t|
  t.libs << 'test'
  t.verbose = true
  t.warning = true
  t.test_files = FileList['test/test_ole.rb']
end

task :default => :test
