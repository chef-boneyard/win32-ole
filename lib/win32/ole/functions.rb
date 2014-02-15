require 'ffi'

module Windows
  module Functions
    extend FFI::Library

    typedef :long, :hresult
    typedef :ulong, :dword

    ffi_lib :ole32

    attach_function :CLSIDFromProgID, [:buffer_in, :pointer], :hresult
    attach_function :CLSIDFromString, [:buffer_in, :pointer], :hresult
    attach_function :CoCreateInstance, [:pointer, :pointer, :dword, :pointer, :pointer], :hresult
    attach_function :CoInitialize, [:pointer], :hresult
    attach_function :CreateBindCtx, [:dword, :pointer], :hresult
    attach_function :MkParseDisplayName, [:pointer, :buffer_in, :pointer, :pointer], :hresult
    attach_function :OleInitialize, [:pointer], :hresult

    ffi_lib :oleaut32

    attach_function :GetActiveObject, [:pointer, :pointer, :pointer], :hresult

    # Macros

    def FAILED(status)
      status < 0
    end

    def SUCCEEDED(status)
      status >= 0
    end
  end
end
