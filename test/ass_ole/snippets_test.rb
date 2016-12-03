require 'test_helper'
module AssOleTest
  describe AssOle::Snippets::LikeOleRuntime do
    describe 'Module like_ole_runtime' do
      runtime = Module.new do
        is_ole_runtime :external
      end

      module_ = Module.new do
        like_ole_runtime runtime
        ole_runtime_get.run Tmp::INFO_BASE
      end

      it 'Module quacks like Ole runtime' do
        module_.sTring('HELLO').must_equal 'HELLO'
      end
    end

    describe 'Class like_ole_runtime' do
      runtime = Module.new do
        is_ole_runtime :external
      end

      klass = Class.new do
        like_ole_runtime runtime
        ole_runtime_get.run Tmp::INFO_BASE
      end

      it 'Class instance quacks like Ole runtime' do
        klass.new.sTring('HELLO').must_equal 'HELLO'
      end

      it 'Class doesn\'t quacks like Ole runtime' do
        proc {
          klass.sTring 'HELLO'
        }.must_raise NoMethodError
      end
    end
  end

  describe AssOle::Snippets::IsSnippet do
    describe 'Ole snippet is Module and mixin for other' do
      snippet = Module.new do
        is_ole_snippet

        def hello_ole(str)
          sTring(str)
        end
      end

      runtime = Module.new do
        is_ole_runtime :external
      end

      itHasRuntime = Module.new do
        it_has_ole_runtime runtime
        extend snippet
      end

      runtime.run Tmp::INFO_BASE

      it 'fail without ole runtime' do
        proc {
          snippet.sTring('HELLO')
        }.must_raise NoMethodError
      end

      it 'sucsess if Snippet mix in to ItHasRuntime' do
        itHasRuntime.hello_ole('HELLO').must_equal('HELLO')
      end
    end

    describe 'Class isn\'t ole snippet' do
      it 'fail' do
        e = proc {
          Class.new do
            is_ole_snippet
          end
        }.must_raise RuntimeError
        e.message.must_match %r{not a Class}
      end
    end

    describe 'Error occurred if ole runtime wasn\'t running' do
      runtime = Module.new do
        is_ole_runtime :external
      end

      module_ = Module.new do
        like_ole_runtime runtime
      end

      # Fail without: runtime.run Tmp::INFO_BASE

      it 'fail without running ole runtime' do
        proc {
          module_.sTring('HELLO')
        }.must_raise AssOle::Snippets::ContextError
      end
    end
  end
end
