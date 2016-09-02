require "json_report_builder/excel_report/object/sheet_object"

module JsonReportBuilder::ExcelReport::Object
  class ExcelObject < Base
    attribute :template_file_name, String
    attribute :password,           String
    attribute :sheets,             Array[SheetObject]
  end
end
