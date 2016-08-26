describe JsonReportBuilder::ExcelReport::Object::ExcelObject do
  context 'camelizeのjson形式で取得できるか' do
    let(:json) { JsonReportBuilder::ExcelReport::Object::ExcelObject.new.to_json_camelize_all_keys }
    context 'Excel Objectのチェック' do
      it 'template_file_nameがtemplateFileNameになっていること' do
        expect(json.include?('templateFileName')).to be_truthy
      end
      it 'passwordがpasswordになっていること' do
        expect(json.include?('password')).to be_truthy
      end
      it 'sheetsがsheetsになっていること' do
        expect(json.include?('sheets')).to be_truthy
      end
    end
  end
end
