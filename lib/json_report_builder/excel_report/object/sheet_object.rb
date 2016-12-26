require 'json_report_builder/excel_report/object/base'
require 'json_report_builder/excel_report/object/row_object'
require 'json_report_builder/excel_report/object/merge_object'

module JsonReportBuilder
  module ExcelReport
    module Object
      class SheetObject < Base
        attribute :template_sheet_name,   String
        attribute :output_sheet_name,     String
        attribute :new_name,              String, default: ''
        attribute :print_col_index_start, Integer, default: -1
        attribute :print_col_index_end,   Integer, default: -1
        attribute :print_row_index_start, Integer, default: -1
        attribute :print_row_index_end,   Integer, default: -1
        attribute :rows,                  Array[RowObject]
        attribute :merges,                Array[MergeObject]

        def create_row(args)
          args = args.merge(row_index: rows.size) if args[:row_index].blank?
          rows << (row = RowObject.new(args))
          row
        end
      end
    end
  end
end