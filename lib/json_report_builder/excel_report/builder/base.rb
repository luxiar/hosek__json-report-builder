require 'virtus'
require 'json_report_builder/excel_report/object/excel_object'

module JsonReportBuilder
  module ExcelReport
    module Builder
      class Base
        include Virtus.model

        attribute :excel_object, Object::ExcelObject

        def initialize
          @excel_object = Object::ExcelObject.new
        end
      end
    end
  end
end
