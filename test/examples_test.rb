require 'test_helper'
module AssOleTest
  describe AssOle::Runtimes do
    describe 'Define runtime' do
      # Specifieis 1C:Enterprise app runtime like :external connection
      external_app = Module.new do
        is_ole_runtime :external
      end

      # Run runtime without fail
      before do
        external_app.run Tmp::INFO_BASE
      end

      it 'runtime is run' do
        external_app.runned?.must_equal true
      end

      # Optional stop runtime or all runtimes will be stopped in at_exit
      after do
        external_app.stop
      end
    end

    describe 'Runtime types' do
      describe ':exeternal connection runtime' do
        external = Module.new do
          is_ole_runtime :external
        end

        it AssOle::Runtimes::App::External do
          external.ole_type.must_equal :external
        end

        it '#run require info_base argument' do
          external.run Tmp::INFO_BASE
        end
      end

      describe ':thick client runtime' do
        thick = Module.new do
          is_ole_runtime :thick
        end

        it AssOle::Runtimes::App::Thick do
          thick.ole_type.must_equal :thick
        end

        it '#run require info_base argument' do
          thick.run Tmp::INFO_BASE
        end
      end

      describe ':thin client runtime' do
        thin = Module.new do
          is_ole_runtime :thin
        end

        it AssOle::Runtimes::App::Thin do
          thin.ole_type.must_equal :thin
        end

        it '#run require info_base argument' do
          thin.run Tmp::INFO_BASE
        end
      end

      describe ':wp 1C:Enterprise server working process connection' do
        wp = Module.new do
          is_ole_runtime :wp
        end

        it AssOle::Runtimes::Claster::Wp do
          wp.ole_class.must_equal AssLauncher::Enterprise::Ole::WpConnection
        end

        it '#run require host:port, platform_require arguments' do
          # It fail because invalid host:port
          proc {
            wp.run 'host:port', '~> 8.3.9.0'
          }.must_raise WIN32OLERuntimeError
        end
      end

      describe ':agent 1C:Enterprise server agent process connection' do
        agent = Module.new do
          is_ole_runtime :agent
        end

        it AssOle::Runtimes::Claster::Agent do
          agent.ole_class.must_equal AssLauncher::Enterprise::Ole::AgentConnection
        end

        it '#run require host:port, platform_require arguments' do
          # It fail because invalid host:port
          proc {
            agent.run 'host:port', '~> 8.3.9.0'
          }.must_raise WIN32OLERuntimeError
        end
      end

      it 'Invalid runtime type' do
        e = proc {
          Module.new do
            is_ole_runtime :bad_type
          end
        }.must_raise RuntimeError
        e.message.must_match %r{Invalid runtime}i
      end
    end

    describe 'For Class Ole runtime will be inharited' do
      parent = Class.new do
        it_has_ole_runtime TEST_OLE_RUNTIME
      end

      child = Class.new(parent)

      it '#ole_connector equal' do
        refute_nil parent.ole_connector
        parent.ole_connector.must_equal\
          child.ole_connector
      end
    end
  end

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

      it_has_ole_runtime = Module.new do
        it_has_ole_runtime runtime
        extend snippet
      end

      runtime.run Tmp::INFO_BASE

      it 'fail without ole runtime' do
        proc {
          snippet.sTring('HELLO')
        }.must_raise NoMethodError
      end

      it 'sucsess if Snippet mix in to it_has_ole_runtime' do
        it_has_ole_runtime.hello_ole('HELLO').must_equal('HELLO')
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

  describe 'IsSnippet and LikeOleRuntime hase helper methods' do
    describe AssOle::Snippets::IsSnippet::WinPath do
      module_ = Module.new do
        like_ole_runtime TEST_OLE_RUNTIME
      end

      # Helper for pass into 1C ole runtime understandable file system pathes.
      it '#real_win_path helper' do
        module_.real_win_path('/').must_match %r{[A-Z]:\\(.)?}i
      end
    end

    describe AssOle::Snippets::IsSnippet::Argv do
      module_ = Module.new do
        like_ole_runtime TEST_OLE_RUNTIME
      end

      # Helper wrapper ower WIN32O::ARGV array
      it '#argv helper' do
        refute_nil module_.argv(0)
      end
    end
  end

  describe 'Use cases' do
    describe 'Write snippets wrappers over the long and awkward 1C syntax' do
      module_ = Module.new do
        it_has_ole_runtime TEST_OLE_RUNTIME
      end

      it 'Long 1C syntax' do
        arr = module_.ole_connector.newObject 'Array'
        arr.add 1
        arr.add 2
        arr.add 3
        # ... etc
        arr.Count.must_equal 3
      end

      # Write snippet
      array_snippet = Module.new do
        is_ole_snippet
        def new_array(*args)
          r = newObject('Array')
          args.each do |a|
            r.add a
          end
          r
        end
      end

      it 'And use short #new_array' do
        module_.send(:extend, array_snippet)
        arr = module_.new_array 1, 2, 3, 4 # ... etc
        arr.Count.must_equal 4
      end
    end

    describe 'Make Ruby script short' do
      long_module = Module.new do
        it_has_ole_runtime TEST_OLE_RUNTIME

        def self.new_array
          ole_connector.newObject 'Array'
        end
      end

      short_module = Module.new do
        like_ole_runtime TEST_OLE_RUNTIME

        def self.new_array
          newObject 'Array'
        end
       end

      it 'Smoky test result' do
        short_module.new_array
        long_module.new_array
      end
    end
  end
end
