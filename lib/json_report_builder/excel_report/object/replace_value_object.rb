require 'json_report_builder/excel_report/object/base'

module JsonReportBuilder
  module ExcelReport
    module Object
      class ReplaceValueObject < Base
        attribute :value,        String
        attribute :replace_text, String
        attribute :direct,       Boolean, default: true
      end
    end
  end
end
