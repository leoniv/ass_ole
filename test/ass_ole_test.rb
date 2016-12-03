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

    fail if runtime.runned?
    runtime.run Tmp::INFO_BASE

    attr_reader :inst
    before do
      @inst = runtimed_klass.new
    end

    it "#{runtimed_klass} includes modues" do
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
      extend AssOle::Runtimes::Claster::Wp
    end

    runtimed_klass = Class.new do
      include runtime
    end

    runtime.run 'host:port', AssOleTest::PLATFORM_REQUIRE
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

    snippet_class = Class.new do
      include runtime
      include AssOle::Snippets::IsSnippet
    end

    attr_reader :inst
    before do
      @inst = snippet_class.new
    end

    it '#{snippited_class} transparent call 1C Ole' do
      inst.sTring('FIXME').must_equal 'FIXME'
    end
  end
end
