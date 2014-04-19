require 'ffi'

module Windows
  module Structs
    extend FFI::Library

    typedef :uchar, :byte

    class GUID < FFI::Struct
      layout(
        :Data1, :ulong,
        :Data2, :ushort,
        :Data3, :ushort,
        :Data4, [:byte, 8]
      )
    end
  end
end
