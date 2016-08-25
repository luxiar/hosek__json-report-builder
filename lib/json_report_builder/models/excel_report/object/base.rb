require 'virtus'

module JsonReportBuilder::ExcelReport::Object
  class Base
    include Virtus.model

    # json形式で値を再帰的に取得　keyは先頭小文字のcamelize 例：keyを template_file_name → templateFileName にする
    def to_json_camelize_all_keys
      values = to_hash.map do |key, value|
        [key.to_s.camelize(:lower), value.is_a?(Array) ? camelize_keys(value) : value]
      end
      Hash[values].to_json
    end

    private

    def camelize_keys(array)
      array.each_with_object([]) do |val, result|
        vals = val.to_hash.map do |key, value|
          [key.to_s.camelize(:lower), value.is_a?(Array) ? camelize_keys(value) : value]
        end
        result << Hash[vals]
      end
    end
  end
end
