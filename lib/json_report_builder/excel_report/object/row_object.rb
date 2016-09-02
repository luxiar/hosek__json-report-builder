require "json_report_builder/excel_report/object/col_object"

module JsonReportBuilder::ExcelReport::Object
  class RowObject < Base
    attribute :row_index,          Integer
    attribute :row_index_template, Integer
    attribute :height_in_points,   Float
    attribute :row_break,          Boolean, default: false
    attribute :cols,               Array[ColObject]
  end
end
