require 'virtus'
require "json_report_builder/models/excel_report/object/excel_object"

module JsonReportBuilder::ExcelReport::Builder
  class Base
    include Virtus.model
    include JsonReportBuilder::ExcelReport::Object
    # ここはvirtusモデルから参照するようで、名前空間を頭から指定しないとattributeが作成できない
    attribute :excel_object, ExcelObject

    def initialize
      @excel_object = ExcelObject.new
    end
  end
end
