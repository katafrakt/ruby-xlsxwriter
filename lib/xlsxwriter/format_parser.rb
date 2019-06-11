module XlsxWriter
  class FormatParser
    attr_reader :workbook

    def initialize(workbook)
      @workbook = workbook
    end

    def format(format)
      format_ptr = C.workbook_add_format(workbook)

      C.format_set_font_size(format_ptr, format[:font_size].to_f) if format[:font_size]
      C.format_set_bold(format_ptr) if format[:bold]
      format_ptr
    end
  end
end
