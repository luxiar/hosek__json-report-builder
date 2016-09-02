require "json_report_builder/version"
require "json_report_builder/config"
require "json_report_builder/excel_report/factory"
require "json_report_builder/excel_report/object/base"
# gem化したときにrequireしなくて使いたいものはここでrequireしておかないといけないみたい
require "json_report_builder/excel_report/builder/base"

module JsonReportBuilder
  include JsonReportBuilder::ExcelReport

  # return output_file_name
  def self.create(tmp_file_name:, builder:, separate: '')
    Factory.new(excel_object: builder.excel_object, separate: separate).create(tmp_file_name)
  end
end
