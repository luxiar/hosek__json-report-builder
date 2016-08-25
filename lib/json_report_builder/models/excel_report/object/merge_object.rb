require "json_report_builder/models/excel_report/object/base"

module JsonReportBuilder::ExcelReport::Object
  class MergeObject < Base
    attribute :start_row_index, Integer
    attribute :start_col_index, Integer
    attribute :end_row_index,   Integer
    attribute :end_col_index,   Integer
  end
end
