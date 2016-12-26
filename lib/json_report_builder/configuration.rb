require 'active_support/configurable'

module JsonReportBuilder
  class Configuration
    include ActiveSupport::Configurable

    config_accessor :tmp_path # 一時的にファイルを作成する場所
    config_accessor :delete_json
  end

  def self.configure
    yield @config ||= JsonReportBuilder::Configuration.new
  end

  def self.config
    @config
  end

  configure do |config|
    config.tmp_path = File.expand_path('tmp') # defaultはprojectのtmpフォルダ
    config.delete_json = true # 一時フォルダに作成したjsonファイルを消すかどうか
  end
end
