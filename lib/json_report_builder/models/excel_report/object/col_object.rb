require "json_report_builder/models/excel_report/object/base"

module JsonReportBuilder::ExcelReport::Object
  class ColObject < Base
    TYPE_IMAGE            = 'Image'.freeze
    TYPE_BOOLEAN          = 'Boolean'.freeze
    TYPE_CALENDAR         = 'Calendar'.freeze
    TYPE_DATE             = 'Date'.freeze
    TYPE_DOUBLE           = 'Double'.freeze
    TYPE_FORMULA          = 'Formula'.freeze
    TYPE_RICH_TEXT_STRING = 'RichTextString'.freeze
    TYPE_STRING           = 'String'.freeze

    attribute :col_index,       Integer
    attribute :value,           String
    attribute :type,            String, default: TYPE_STRING
    # TYPE_IMAGEのときに、画像サイズをそのまま使いたい場合に-1を指定する
    attribute :pic_row_index_e, Integer, default: -1
    attribute :pic_col_index_e, Integer, default: -1
  end
end
