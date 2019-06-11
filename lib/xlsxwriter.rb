require 'xlsxwriter/version'
require 'xlsxwriter/c'
require 'xlsxwriter/format_parser'

module XlsxWriter
  class Error < StandardError; end
  # Your code goes here...

  def self.create(filename, options = {})
    wbook = Workbook.new(filename, options)
    yield(wbook)
  ensure
    wbook.close if wbook
  end

  class Workbook
    def initialize(filename, options = {})
      workbook_opts = parse_options(options)
      @workbook = C.workbook_new_opt(filename, workbook_opts)
      # TODO: support options
      @worksheet = C.workbook_add_worksheet(workbook, nil)
      @current_row = 0
      @format_parser = FormatParser.new(workbook)
    end

    def close
      C.workbook_close(workbook)
    end

    def write_row(rowdef, format = {})
      offset = format.delete(:offset) || 0
      apply_row_formatting!(format)

      rowdef.each_with_index do |item, idx|
        next if item.nil?

        idx += offset
        if item.is_a?(Numeric)
          C.worksheet_write_number(worksheet, current_row, idx, item, nil)
        else
          C.worksheet_write_string(worksheet, current_row, idx, item.to_s, nil)
        end
      end
      @current_row += 1
    end

    def skip_row
      @current_row += 1
    end

    private

    attr_reader :workbook, :worksheet, :current_row, :format_parser

    def parse_options(options)
      C::Options.new.tap do |opts|
        const_memory = options[:constant_memory] ? C::LXW_BOOLEAN[:lxw_true] : C::LXW_BOOLEAN[:lxw_false]
        opts[:constant_memory] = const_memory
      end
    end

    def apply_row_formatting!(format)
      return if format == {}

      height = format.fetch(:height, C::LXW_DEF_ROW_HEIGHT)
      format_ptr = format_parser.format(format)
      C.worksheet_set_row(worksheet, current_row, height, format_ptr)
    end
  end
end
