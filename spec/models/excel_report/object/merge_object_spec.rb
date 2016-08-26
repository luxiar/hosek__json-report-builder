describe JsonReportBuilder::ExcelReport::Object::MergeObject do
  context 'camelizeのjson形式で取得できるか' do
    let(:json) { JsonReportBuilder::ExcelReport::Object::MergeObject.new.to_json_camelize_all_keys }
    context 'Merge Objectのチェック' do
      it 'start_row_indexがstartRowIndexになっていること' do
        expect(json.include?('startRowIndex')).to be_truthy
      end
      it 'start_col_indexがstartColIndexになっていること' do
        expect(json.include?('startColIndex')).to be_truthy
      end
      it 'end_row_indexがendRowIndexになっていること' do
        expect(json.include?('endRowIndex')).to be_truthy
      end
      it 'end_col_indexがendColIndexになっていること' do
        expect(json.include?('endColIndex')).to be_truthy
      end
    end
  end
end
