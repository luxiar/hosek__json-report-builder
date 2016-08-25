require 'active_support/configurable'
module JsonReportBuilder
  def self.configure(&block)
    yield @config ||= JsonReportBuilder::Configuration.new
  end

  def self.config
    @config
  end

  class Configuration
    include ActiveSupport::Configurable
    config_accessor :tmp_path  # 一時的にファイルを作成する場所
    config_accessor :delete_json
  end

  configure do |config|
    config.tmp_path = File.expand_path('tmp')  # defaultはprojectのtmpフォルダ
    config.delete_json = true # 一時フォルダに作成したjsonファイルを消すかどうか
  end
end
