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
        :Data4, [:uchar, 8]   # First 2 bytes contain the 3rd group of 4 hex digits. Remaining 6 bytes contain the final 12 hex digits.
      )

      def initialize(long = nil, short1 = nil, short2 = nil, *bytes)
        super()
        self[:Data1] = long if long
        self[:Data2] = short1 if short1
        self[:Data3] = short2 if short2
        self[:Data4] = bytes.pack("C*") if !bytes.empty?
      end

      def vtable
      end

      def query_interface
      end
    end
  end
end
