require 'json_report_builder/excel_report/object/base'

module JsonReportBuilder
  module ExcelReport
    module Object
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
        # mergeしたcellの場合、自動高さ調節ができないので、計算で高さ調節してあげる
        # その場合に、列数を基準とするので、mergeした範囲の列数を指定する（ただしぴったりフィットはしません）
        attribute :calc_height_for_merge_col, Integer
        # 動的に列が増える場合に、コピー元となるtemplateシートのセルを指定する
        attribute :copy_row_index, Integer, default: -1
        attribute :copy_col_index, Integer, default: -1
      end
    end
  end
end
