require 'rake'
require 'rake/testtask'
require 'rbconfig'
include Config

desc 'Install the win32-ole library (non-gem)'
task :install do
   sitelibdir = CONFIG['sitelibdir']
   installdir = File.join(sitelibdir, 'win32')
   file = 'lib\win32\ole.rb'

   Dir.mkdir(installdir) unless File.exists?(installdir)
   FileUtils.cp(file, installdir, :verbose => true)
end

Rake::TestTask.new do |t|
   t.libs << 'test'
   t.verbose = true
   t.warning = true
   t.test_files = FileList['test/test_ole.rb']
end
