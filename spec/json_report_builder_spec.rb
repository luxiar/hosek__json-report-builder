require 'spec_helper'
require 'models/excel_report/builder/sample_builder'

describe JsonReportBuilder do
  it 'has a version number' do
    # expect(JsonReportBuilder::VERSION).not_to be nil
  end

  describe '出力実行' do
    subject { File.exist?(output_file_name) }

    let(:builder) { JsonReportBuilder::ExcelReport::Builder::SampleBuilder.new }
    let(:output_file_name) { JsonReportBuilder.create(tmp_file_name: 'test', builder: builder, separate: 'sep') }

    # 作成したエクセルファイルが見たい場合には、afterをコメント化
    before { builder.sample }
    after { File.delete(output_file_name) if File.exist?(output_file_name) }

    context 'templateファイルが見つかった場合' do
      it '正常にエクセルファイルが作られること' do
        is_expected.to be_truthy
      end
    end
    context 'templateファイルが見つからない場合' do
      before { builder.excel_object.template_file_name = '' }
      it 'エクセルファイルが作成されないこと' do
        is_expected.to be_falsy
      end
    end
  end
end
