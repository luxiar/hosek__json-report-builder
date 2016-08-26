require 'rails'
require 'fileutils'

module JsonReportBuilder::ExcelReport
  class Factory
    def initialize(excel_object:, separate: '')
      @excel_object = excel_object
      @separate = separate
    end

    #
    # 帳票を作成し、できた帳票のファイルパスを返します
    #
    def create(file_name)
      create_params(file_name)
      command = create_command_str
      result = system(command)
      fail '帳票の作成に失敗しました' unless result

      File.delete(@json_file_name) if JsonReportBuilder.config.delete_json && FileTest.exist?(@json_file_name)
      @output_file_name
    end

    private

    def create_params(file_name)
      @json_file_name = json_excel_reports(file_name)
      File.open(@json_file_name, 'w') do |f|
        f.write(@excel_object.to_json_camelize_all_keys)
      end
      @output_file_name = output_file(file_name)
    end

    def create_command_str
      command_strs = []
      command_strs << 'java -jar'
      command_strs << File.join(File.expand_path('../../../', __FILE__), 'jar', 'json_report_builder.jar')
      command_strs << @json_file_name
      command_strs << @output_file_name
      command_strs.join(' ')
    end

    def json_excel_reports(file_name)
      path = File.join(JsonReportBuilder.config.tmp_path, @separate)
      FileUtils.mkdir_p(path)
      "#{ path + file_name }_#{ Time.current.strftime('%Y%m%d-%H%M%S') }.json"
    end

    def output_file(file_name)
      path = File.join(JsonReportBuilder.config.tmp_path, @separate)
      FileUtils.mkdir_p(path)
      "#{ path + file_name }_#{ Time.current.strftime('%Y%m%d-%H%M%S') }.xlsx"
    end
  end
end
