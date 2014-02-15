require 'ffi'

module Windows
  module Functions
    extend FFI::Library

    typedef :hresult, :long

    ffi_lib :ole32

    attach_function :CLSIDFromProgID, [:buffer_in, :pointer], :hresult
    attach_function :CLSIDFromString, [:buffer_in, :pointer], :hresult
    attach_function :CoCreateInstance, [:pointer, :pointer, :dword, :pointer, :pointer], :hresult
    attach_function :CoInitialize, [:pointer], :hresult

    # Macros

    def FAILED(status)
      status < 0
    end

    def SUCCEEDED(status)
      status >= 0
    end
  end
end
