require 'test_helper'
module AssOleTest
  describe AssOle::Snippets::IsSnippet do
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
        inst.string_get1('HELLO').must_equal 'HELLO'
        inst.string_get2('HELLO').must_equal 'HELLO'
        inst.sTring('HELLO').must_equal 'HELLO'
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
        inst.sTring('HELLO').must_equal 'HELLO'
      end
    end

    describe 'Module self is snippet' do
      runtime = Module.new do
        extend AssOle::Runtimes::App::External
      end

      runtime.run Tmp::INFO_BASE

      snippet_module = Module.new do
        extend runtime
        extend AssOle::Snippets::IsSnippet
      end

      it '#{snippet_module} transparent call 1C Ole' do
        snippet_module.sTring('HELLO').must_equal 'HELLO'
      end
    end
  end
end
