require 'ffi'

module Windows
  module Structs
    extend FFI::Library

    class GUID < FFI::Struct
      layout(
        :Data1, :ulong,
        :Data2, :ushort,
        :Data3, :ushort,
        :Data4, [:char, 8]
      )
    end
  end
end
