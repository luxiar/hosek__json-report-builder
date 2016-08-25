require 'spec_helper'
require "json_report_builder/models/excel_report/object/sheet_object"

describe JsonReportBuilder::ExcelReport::Object::SheetObject do
  context 'camelizeのjson形式で取得できるか' do
    let(:json) { JsonReportBuilder::ExcelReport::Object::SheetObject.new.to_json_camelize_all_keys }
    context 'Sheet Objectのチェック' do
      it 'template_sheet_nameがtemplateSheetNameになっていること' do
        expect(json.include?('templateSheetName')).to be_truthy
      end
      it 'output_sheet_nameがoutputSheetNameになっていること' do
        expect(json.include?('outputSheetName')).to be_truthy
      end
      it 'rowsがrowsになっていること' do
        expect(json.include?('rows')).to be_truthy
      end
      it 'mergesがmergesになっていること' do
        expect(json.include?('merges')).to be_truthy
      end
    end
  end
end
