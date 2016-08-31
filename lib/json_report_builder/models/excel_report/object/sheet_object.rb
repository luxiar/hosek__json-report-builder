require "json_report_builder/models/excel_report/object/row_object"
require "json_report_builder/models/excel_report/object/merge_object"

module JsonReportBuilder::ExcelReport::Object
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
  end
end
