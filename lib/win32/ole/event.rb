module Win32
  class OLE
    # Just a stub for now.
    class Event
      def initialize
      end

      # I'll probably change these method names.
      def on_event(*args)
        yield
      end

      def on_event_with_outargs(*args)
        yield
      end

      def off_event(arg=nil)
      end

      def unadvise
      end

      def handler
      end

      def handler=(val)
      end

      def message_loop
      end
    end
  end
end
