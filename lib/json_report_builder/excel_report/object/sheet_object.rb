require 'json_report_builder/excel_report/object/base'
require 'json_report_builder/excel_report/object/row_object'
require 'json_report_builder/excel_report/object/merge_object'
require 'json_report_builder/excel_report/object/replace_value_object'

module JsonReportBuilder
  module ExcelReport
    module Object
      class SheetObject < Base
        attribute :template_sheet_name,   String
        attribute :output_sheet_name,     String
        attribute :output_sheet_hide,     Boolean, default: false
        attribute :new_name,              String,  default: ''
        attribute :copy_output_sheet,     Boolean, default: false
        attribute :print_col_index_start, Integer, default: -1
        attribute :print_col_index_end,   Integer, default: -1
        attribute :print_row_index_start, Integer, default: -1
        attribute :print_row_index_end,   Integer, default: -1
        attribute :rows,                  Array[RowObject]
        attribute :merges,                Array[MergeObject]
        attribute :replace_values,        Array[ReplaceValueObject]

        def create_row(args)
          args = args.merge(row_index: rows.size) if args[:row_index].blank?
          rows << (row = RowObject.new(args))
          row
        end

        def create_replace_value(args)
          args[:value] = '' if args[:value].nil? && !args[:use_nil]
          replace_values << (replace_value = ReplaceValueObject.new(args))
          replace_value
        end
      end
    end
  end
end
