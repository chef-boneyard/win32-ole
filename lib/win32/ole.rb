require 'socket'
require File.join(File.dirname(__FILE__), 'ole', 'constants')
require File.join(File.dirname(__FILE__), 'ole', 'structs')
require File.join(File.dirname(__FILE__), 'ole', 'functions')
require File.join(File.dirname(__FILE__), 'ole', 'helper')

module Win32
  class OLE
    include Windows::Constants
    include Windows::Structs
    include Windows::Functions
    extend Windows::Functions

    # The version of the win32-ole library.
    VERSION = '0.1.0'

    # The name of the OLE automation server specified in the constructor.
    attr_reader :server

    # The host the OLE automation object was created on.
    attr_reader :host

    # Interface GUID's.

    IID_IUnknown     = [0,0,0,192,0,0,0,0,0,0,70].pack('ISSCCCCCCCC')
    IID_IDispatch    = [132096,0,0,192,0,0,0,0,0,0,70].pack('ISSCCCCCCCC')
    IID_IEnumVARIANT = [132100,0,0,192,0,0,0,0,0,0,70].pack('ISSCCCCCCCC')

    # Creates a new Win32::OLE server object on +host+, or the localhost if
    # no host is specified. The +server+ can be either a Program ID or a
    # Class ID.
    #
    # Examples:
    #
    #    # Program ID (Excel)
    #    ole = Win32::OLE.new('Excel.Application')
    #
    #    # Class ID (Excel)
    #    ole = Win32::OLE.new('{00024500-0000-0000-C000-000000000046}')
    #
    def initialize(server, host = Socket.gethostname)
      raise TypeError unless server.is_a?(String)
      raise TypeError unless host.is_a?(String)

      @server = server
      @host   = host

      clsid = FFI::MemoryPointer.new(:char, 16)

      # Attempt to get a CLSID using from both ProgID and String
      hr = CLSIDFromProgID(server.wincode, clsid)

      if FAILED(hr)
        hr = CLSIDFromString(server.wincode, clsid)

        if FAILED(hr)
          FFI.raise_windows_error('CLSIDFromString')
        end
      end

      hr = CoInitialize(nil)

      FFI.raise_windows_error('CoInitialize') if FAILED(hr)

      dispatch = FFI::MemoryPointer.new(:char, IID_IDispatch.size)

      hr = CoCreateInstance(
        clsid,
        nil,
        CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
        IID_IDispatch,
        dispatch
      )

      FFI.raise_windows_error('CoCreateInstance') if FAILED(hr)

      @dispatch = dispatch
    end

    # Opens an existing OLE server object on +host+, or the localhost if
    # no host is specified. The +server+ can be a Program ID, a Class ID,
    # or a moniker.
    #
    # Examples:
    #
    #    # Program ID (Excel)
    #    ole = Win32::OLE.open('Excel.Application')
    #
    #    # Class ID (Excel)
    #    ole = Win32::OLE.open('{00024500-0000-0000-C000-000000000046}')
    #
    #    # Moniker
    #    ole = Win32::OLE.open('winmgmts://some_host/root/cimv2')
    #
    def self.open(server, host=Socket.gethostname)
      server = server.wincode
      hr = OleInitialize(nil)

      FFI.raise_windows_error('OleInitialize') if FAILED(hr)

      clsid = FFI::MemoryPointer.new(:char, 16)

      # Try as a Program ID first
      hr = CLSIDFromProgID(server, clsid)

      # Then try as a Class ID
      if FAILED(hr)
        hr = CLSIDFromString(server, clsid)

        # Finally, try as a moniker
        if FAILED(hr)
          ctx = FFI::MemoryPointer.new(:ulong)
          hr  = CreateBindCtx(0, ctx)

          FFI.raise_windows_error('CreateBindCtx') if FAILED(hr)

          eaten   = FFI::MemoryPointer.new(:ulong)
          moniker = FFI::MemoryPointer.new(:ulong)

          hr = MkParseDisplayName(ctx, server, eaten, moniker)

          FFI.raise_windows_error('MkParseDisplayName') if FAILED(hr)

          # TODO: Now what?

          return
        end
      end

      pptr = FFI::MemoryPointer.new(GUID)

      hr = GetActiveObject(clsid, nil, pptr)

      FFI.raise_windows_error('GetActiveObject') if FAILED(hr)

      iunknown = GUID.new(pptr.read_pointer)

      # TODO: Now what?
    end
  end
end

if $0 == __FILE__
  #excel = Win32::OLE.open('Excel.Application')
  #my_comp = Win32::OLE.open('{20d04fe0-3aea-1069-a2d8-08002b30309d}')
  #ie = Win32::OLE.new('InternetExplorer.Application')
  ie = Win32::OLE.open('InternetExplorer.Application')
  p ie
end
