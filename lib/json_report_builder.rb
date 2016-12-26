require 'json_report_builder/configuration'
require 'json_report_builder/excel_report/builder/base'

module JsonReportBuilder
  autoload :Version, 'json_report_builder/version'

  module ExcelReport
    autoload :Factory, 'json_report_builder/excel_report/factory'

    module Object
      autoload :Base,        'json_report_builder/excel_report/object/base'
      autoload :ColObject,   'json_report_builder/excel_report/object/col_object'
      autoload :ExcelObject, 'json_report_builder/excel_report/object/excel_object'
      autoload :MergeObject, 'json_report_builder/excel_report/object/merge_object'
      autoload :RowObject,   'json_report_builder/excel_report/object/row_object'
      autoload :SheetObject, 'json_report_builder/excel_report/object/sheet_object'
    end
  end

  # return output_file_name
  def self.create(tmp_file_name:, builder:, separate: '')
    ExcelReport::Factory.new(excel_object: builder.excel_object, separate: separate).create(tmp_file_name)
  end
end
