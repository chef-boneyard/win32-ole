#######################################################################
# test_ole.rb
#
# Test suite for the win32-ole library. You should run these tests
# via the 'rake test' task.
#######################################################################
require 'win32/ole'
require 'test/unit'
include Win32

class TC_Win32_OLE < Test::Unit::TestCase
  def setup
    @server = 'InternetExplorer.Application'
    @host   = Socket.gethostname
  end

  def test_constructor
    assert_respond_to(OLE, :new)
    assert_nothing_raised{ OLE.new(@server) }
    assert_nothing_raised{ OLE.new(@server, @host) }
  end

  def test_constructor_expected_errors
    assert_raise(ArgumentError){ OLE.new }
    assert_raise(ArgumentError){ OLE.new(@server, @host, true) }
  end

  def test_constructor_expected_security_errors
    proc do
      $SAFE = 1
      @server.taint
      assert_raise(SecurityError){ OLE.new(@server) }
    end.call

    proc do
      $SAFE = 1
      @host.taint
      server = 'InternetExplorer.Application'
      assert_raise(SecurityError){ OLE.new(server, @host) }
    end.call
  end

  def teardown
    @server = nil
    @host   = nil
  end
end
