require "json_report_builder/version"
require "json_report_builder/config"
require "json_report_builder/models/excel_report/factory"

module JsonReportBuilder
  include JsonReportBuilder::ExcelReport

  # return output_file_name
  def self.build(tmp_file_name:, excel_object:, separate: '')
    Factory.new(excel_object: excel_object, separate: separate).create(tmp_file_name)
  end
end
