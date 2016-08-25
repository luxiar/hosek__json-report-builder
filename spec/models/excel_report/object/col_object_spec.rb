require 'spec_helper'
require "json_report_builder/models/excel_report/object/col_object"

describe JsonReportBuilder::ExcelReport::Object::ColObject do
  context 'camelizeのjson形式で取得できるか' do
    let(:json) { JsonReportBuilder::ExcelReport::Object::ColObject.new.to_json_camelize_all_keys }
    context 'Col Objectのチェック' do
      it 'col_indexがcolIndexになっていること' do
        expect(json.include?('colIndex')).to be_truthy
      end
      it 'valueがvalueになっていること' do
        expect(json.include?('value')).to be_truthy
      end
      it 'typeがtypeになっていること' do
        expect(json.include?('type')).to be_truthy
      end
      it 'pic_row_index_eがpicRowIndexEになっていること' do
        expect(json.include?('picRowIndexE')).to be_truthy
      end
      it 'pic_col_index_eがpicColIndexEになっていること' do
        expect(json.include?('picColIndexE')).to be_truthy
      end
    end
  end
end
