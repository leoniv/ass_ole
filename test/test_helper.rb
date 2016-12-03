$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)
require 'simplecov'
at_exit do
  AssOleTest::Tmp.do_at_exit
end

require 'ass_maintainer/info_base'
require 'ass_ole'
require 'minitest/autorun'

module AssOleTest
  PLATFORM_REQUIRE = '~> 8.3.9.0'
  module Tmp
    extend AssLauncher::Api
    INFO_BASE_PATH = File.join(Dir.tmpdir, 'ass_ole_test.ib')
    INFO_BASE_CS = cs_file file: INFO_BASE_PATH
    INFO_BASE = AssMaintainer::InfoBase
      .new('ass_ole_test', INFO_BASE_CS, false, platform_require: PLATFORM_REQUIRE)
    INFO_BASE.make

    def self.do_at_exit
      INFO_BASE.rm! :yes
    end
  end
end
