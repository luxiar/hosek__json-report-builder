require 'json_report_builder/excel_report/object/base'
require 'json_report_builder/excel_report/object/sheet_object'

module JsonReportBuilder
  module ExcelReport
    module Object
      class ExcelObject < Base
        attribute :template_file_name, String
        attribute :password,           String
        attribute :sheets,             Array[SheetObject]

        def create_sheet(args)
          sheets << (sheet = SheetObject.new(args))
          sheet
        end
      end
    end
  end
end
