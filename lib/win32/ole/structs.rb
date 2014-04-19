require 'ffi'

module Windows
  module Structs
    extend FFI::Library

    typedef :uchar, :byte

    class GUID < FFI::Struct
      layout(
        :Data1, :ulong,       # First 8 hex digits
        :Data2, :ushort,      # First group of 4 hex digits
        :Data3, :ushort,      # Second group of 4 hex digits
        :Data4, [:byte, 8]    # First 2 bytes contain the 3rd group of 4 hex digits. Remaining 6 bytes contain the # final 12 hex digits.
      )
    end
  end
end
