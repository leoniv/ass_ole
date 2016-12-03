require 'test_helper'

module AssOleTest
  describe ::AssOle::VERSION do
    it 'has a version number' do
      refute_nil ::AssOle::VERSION
    end
  end

  describe AssOle::Runtimes::App::External do
    runtime = Module.new do
      extend AssOle::Runtimes::App::External
    end

    runtimed_klass = Class.new do
      include runtime
    end

    runtimed_module = Module.new do
      extend runtime
    end

    fail if runtime.runned?
    runtime.run Tmp::INFO_BASE

    it 'runtimes equals' do
      runtimed_klass.ole_connector.must_equal\
        runtimed_module.ole_connector
    end

    attr_reader :inst
    before do
      @inst = runtimed_klass.new
    end

    it '.ole_connector' do
      runtimed_klass.ole_connector.must_be_instance_of\
        AssLauncher::Enterprise::Ole::IbConnection
      runtimed_module.ole_connector.must_be_instance_of\
        AssLauncher::Enterprise::Ole::IbConnection
    end

    it '#ole_connector' do
      inst.ole_connector.must_be_instance_of\
        AssLauncher::Enterprise::Ole::IbConnection
    end
  end

  describe AssOle::Runtimes::App::Thick do
    runtime = Module.new do
      extend AssOle::Runtimes::App::Thick
    end

    runtime.run Tmp::INFO_BASE

    it '.ole_connector' do
      runtime.ole_connector.must_be_instance_of\
        AssLauncher::Enterprise::Ole::ThickApplication
    end
  end

  describe AssOle::Runtimes::App::Thin do
    runtime = Module.new do
      extend AssOle::Runtimes::App::Thin
    end

    runtime.run Tmp::INFO_BASE

    it '.ole_connector' do
      runtime.ole_connector.must_be_instance_of\
        AssLauncher::Enterprise::Ole::ThinApplication
    end
  end

  describe AssOle::Runtimes::Claster::Agent do
    runtime = Module.new do
      extend AssOle::Runtimes::Claster::Agent
    end

    it 'fail because bad hostname' do
      e = proc {
        runtime.run 'host:port', AssOleTest::PLATFORM_REQUIRE
      }.must_raise WIN32OLERuntimeError
      e.message.force_encoding('ASCII-8BIT').split('HRESULT')[0]
        .force_encoding('UTF-8')
        .must_match %r{line=\d+ file=src\\DataExchangeCommon\.cpp}
    end

    it 'Agent connection' do
      runtime.ole_class.must_equal\
        AssLauncher::Enterprise::Ole::AgentConnection
    end
  end

  describe AssOle::Runtimes::Claster::Wp do
    runtime = Module.new do
      extend AssOle::Runtimes::Claster::Wp
    end

    it 'fail because bad hostname' do
      e = proc {
        runtime.run 'host:port', AssOleTest::PLATFORM_REQUIRE
      }.must_raise WIN32OLERuntimeError
      e.message.force_encoding('ASCII-8BIT').split('HRESULT')[0]
        .force_encoding('UTF-8')
        .must_match %r{line=\d+ file=src\\DataExchangeCommon\.cpp}
    end

    it 'Wp connection' do
      runtime.ole_class.must_equal\
        AssLauncher::Enterprise::Ole::WpConnection
    end
  end
end
