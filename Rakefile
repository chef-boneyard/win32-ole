require 'rake'
require 'rake/testtask'
require 'rake/clean'
require 'rake/extensiontask'
require 'rbconfig'
require 'tmpdir'
include RbConfig

Rake::ExtensionTask.new('ole')

CLEAN.include(
  '**/*.gem',               # Gem files
  '**/*.rbc',               # Rubinius
  '**/*.o',                 # C object file
  '**/*.log',               # Ruby extension build log
  '**/Makefile',            # C Makefile
  '**/*.def',               # Definition files
  '**/*.exp',
  '**/*.lib',
  '**/*.pdb',
  '**/*.obj',
  '**/*.stackdump',         # Junk that can happen on Windows
  "**/*.#{CONFIG['DLEXT']}" # C shared object
)

Rake::TestTask.new do |t|
  t.libs << 'test'
  t.verbose = true
  t.warning = true
  t.test_files = FileList['test/test_ole.rb']
end

task :default => :test
