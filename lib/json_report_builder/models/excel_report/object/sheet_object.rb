require "json_report_builder/models/excel_report/object/base"
require "json_report_builder/models/excel_report/object/row_object"
require "json_report_builder/models/excel_report/object/merge_object"

module JsonReportBuilder::ExcelReport::Object
  class SheetObject < Base
    attribute :template_sheet_name, String
    attribute :output_sheet_name,   String
    attribute :rows,                Array[RowObject]
    attribute :merges,              Array[MergeObject]
  end
end
