require 'json_report_builder/excel_report/object/base'
require 'json_report_builder/excel_report/object/col_object'

module JsonReportBuilder
  module ExcelReport
    module Object
      class ReplaceValueObject < Base
        attribute :value,        String
        attribute :replace_text, String
        attribute :type,         String, default: ColObject::TYPE_STRING
      end
    end
  end
end
