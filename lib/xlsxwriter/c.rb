require 'ffi'

module XlsxWriter
  module C
    extend FFI::Library
    ffi_lib 'xlsxwriter'

    LXW_BOOLEAN = enum(:lxw_boolean, [:lxw_false, :lxw_true])

    class Options < FFI::Struct
      layout :constant_memory, LXW_BOOLEAN
    end

    # TODO: probably better to read it from C code
    LXW_DEF_ROW_HEIGHT = 15.0 # apparently a default, see: https://libxlsxwriter.github.io/worksheet_8h.html#a8901b9706d1c48c28c97e95b452a927a
    LXW_DEF_COL_WIDTH = 8.43

    attach_function :workbook_new_opt, [:string, Options], :pointer
    attach_function :workbook_add_worksheet, [:pointer, :pointer], :pointer
    attach_function :worksheet_write_string, [:pointer, :int, :int, :string, :pointer], :void
    attach_function :worksheet_write_number, [:pointer, :int, :int, :double, :pointer], :void
    attach_function :workbook_close, [:pointer], :void
    attach_function :worksheet_set_row, [:pointer, :uint, :double, :pointer], :void
    attach_function :worksheet_set_column, [:pointer, :uint, :uint, :double, :pointer], :void

    # formatting
    attach_function :workbook_add_format, [:pointer], :pointer
    attach_function :format_set_bold, [:pointer], :void
    attach_function :format_set_font_size, [:pointer, :double], :void
  end
end
