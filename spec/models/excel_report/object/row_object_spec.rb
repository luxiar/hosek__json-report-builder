require 'spec_helper'
require "json_report_builder/models/excel_report/object/row_object"

describe JsonReportBuilder::ExcelReport::Object::RowObject do
  context 'camelizeのjson形式で取得できるか' do
    let(:json) { JsonReportBuilder::ExcelReport::Object::RowObject.new.to_json_camelize_all_keys }
    context 'Row Objectのチェック' do
      it 'row_indexがrowIndexになっていること' do
        expect(json.include?('rowIndex')).to be_truthy
      end
      it 'row_index_templateがrowIndexTemplateになっていること' do
        expect(json.include?('rowIndexTemplate')).to be_truthy
      end
      it 'height_in_pointsがheightInPointsになっていること' do
        expect(json.include?('heightInPoints')).to be_truthy
      end
      it 'colsがcolsになっていること' do
        expect(json.include?('cols')).to be_truthy
      end
    end
  end
end
