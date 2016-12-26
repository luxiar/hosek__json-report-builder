require 'json_report_builder/excel_report/object/base'
require 'json_report_builder/excel_report/object/col_object'

module JsonReportBuilder
  module ExcelReport
    module Object
      class RowObject < Base
        attribute :row_index,          Integer
        attribute :row_index_template, Integer
        attribute :height_in_points,   Float
        attribute :row_break,          Boolean, default: false
        attribute :cols,               Array[ColObject]

        def create_col(args)
          cols << (col = ColObject.new(args))
          col
        end
      end
    end
  end
end
