require 'windows/com'
require 'windows/com/automation'
require 'windows/error'
require 'windows/unicode'
require 'windows/registry'
require 'socket'

module Win32   
   class OLE
      include Windows::Error
      include Windows::COM
      include Windows::COM::Automation
      include Windows::Unicode
      include Windows::Registry

      extend Windows::Error
      extend Windows::COM
      extend Windows::COM::Automation
      extend Windows::Unicode
      extend Windows::Registry
      
      # The version of the win32-ole library.
      VERSION = '0.1.0'

      # Error raised if any of the OLE related methods fail.
      class Error < StandardError; end

      # The name of the OLE automation server specified in the constructor.
      attr_reader :server

      # The host the OLE automation object was created on.
      attr_reader :host

      # These definitions were taken from http://tinyurl.com/3z8z4h
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
         @server = server
         @host   = host

         clsid    = 0.chr * 16
         dispatch = 0.chr * IID_IDispatch.size

         # Attempt to get a CLSID using from both ProgID and String
         hr = CLSIDFromProgID(multi_to_wide(server, CP_UTF8), clsid)

         if FAILED(hr)
            hr = CLSIDFromString(multi_to_wide(server, CP_UTF8), clsid)
            if FAILED(hr)
               raise Error, "unknown OLE server '#{server}'"
            end
         end

         hr = CoInitialize(nil)

         if FAILED(hr)
            raise Error, get_last_error
         end
      
         hr = CoCreateInstance(
            clsid,
            nil,
            CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
            IID_IDispatch,
            dispatch
         )

         if FAILED(hr)
            raise Error, "failed to create OLE object from '#{server}'"
         end

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
         server = multi_to_wide(server, CP_UTF8)    
         hr = OleInitialize()

         if FAILED(hr)
            raise Error, get_last_error
         end

         clsid = 0.chr * 16

         # Try as a Program ID first
         hr = CLSIDFromProgID(server, clsid)

         # Then try as a Class ID
         if FAILED(hr)
            hr = CLSIDFromString(server, clsid)
            
            # Finally, try as a moniker
            if FAILED(hr)
               bind_ctx = 0.chr * 4
               hr = CreateBindCtx(0, bind_ctx)
               
               if FAILED(hr)
                  raise Error, get_last_error
               end
               
               bind_ctx = bind_ctx.unpack('L').first
               eaten   = 0.chr * 4
               moniker = 0.chr * 4

               hr = MkParseDisplayName(bind_ctx, server, eaten, moniker)

               if FAILED(hr)
                  raise Error, get_last_error
               end

               # TODO: Now what?

               return
            end
         end

         unknown = 0.chr * IID_IUnknown.size
         
         hr = GetActiveObject(clsid, 0, unknown)

         if FAILED(hr)
            raise Error, get_last_error
         end

         # TODO: Now what?
      end
   end 
end
