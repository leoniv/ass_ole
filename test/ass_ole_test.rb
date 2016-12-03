require 'test_helper'

module AssOleTest
  describe ::AssOle::VERSION do
    it 'has a version number' do
      refute_nil ::AssOle::VERSION
    end
  end

  describe AssOle::Runtimes::App do
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

    it 'runtime equals' do
      runtimed_klass.ole_connector.must_equal\
        runtimed_module.ole_connector
    end

    attr_reader :inst
    before do
      @inst = runtimed_klass.new
    end

    it "#{runtimed_klass} includes modues" do
      skip
      assert runtimed_klass.include? AssOle::Runtimes::ModuleMethods
      assert runtimed_klass.include? AssOle::Runtimes::RuntimeDispatcher
    end

    it '.ole_connector' do
      runtimed_klass.ole_connector.must_be_instance_of\
        AssLauncher::Enterprise::Ole::IbConnection
    end

    it '#ole_connector' do
      inst.ole_connector.must_be_instance_of\
        AssLauncher::Enterprise::Ole::IbConnection
    end
  end

  describe AssOle::Runtimes::Claster do
    runtime = Module.new do
      extend AssOle::Runtimes::Claster::Agent
    end

    runtimed_klass = Class.new do
      include runtime
    end

    runtimed_module = Module.new do
      extend runtime
    end

    it 'fail because bad hostname' do
      e = proc {
        runtime.run 'host:port', AssOleTest::PLATFORM_REQUIRE
      }.must_raise WIN32OLERuntimeError
      e.message.force_encoding('ASCII-8BIT').split('HRESULT')[0]
        .force_encoding('UTF-8')
        .must_match %r{line=\d+ file=src\\DataExchangeCommon\.cpp}
    end
  end

  describe 'Class includes snippets' do
    snippet1 = Module.new do
      extend AssOle::Snippets::IsSnippet

      def string_get1(obj)
        sTring(obj)
      end
    end

    snippet2 = Module.new do
      extend AssOle::Snippets::IsSnippet

      def string_get2(obj)
        sTring(obj)
      end
    end

    runtime = Module.new do
      extend AssOle::Runtimes::App::External
    end

    snippited_class = Class.new do
      include runtime
      include snippet1
      include snippet2
    end

    runtime.run Tmp::INFO_BASE

    attr_reader :inst
    before do
      @inst = snippited_class.new
    end

    it '#{snippited_class} transparent call 1C Ole' do
      inst.string_get1('FIXME').must_equal 'FIXME'
      inst.string_get2('FIXME').must_equal 'FIXME'
      inst.sTring('FIXME').must_equal 'FIXME'
    end
  end

  describe 'Class self is snippet' do
    runtime = Module.new do
      extend AssOle::Runtimes::App::External
    end

    runtime.run Tmp::INFO_BASE

    snippet_class = Class.new do
      include runtime
      include AssOle::Snippets::IsSnippet
    end

    attr_reader :inst
    before do
      @inst = snippet_class.new
    end

    it '#{snippet_class} transparent call 1C Ole' do
      inst.sTring('FIXME').must_equal 'FIXME'
    end
  end
end
