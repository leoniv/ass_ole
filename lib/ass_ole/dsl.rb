module AssOle
  # It mixins for {Module}
  module DSL
    # Define module as Ole snippet
    # @example
    #  module Query
    #    is_ole_snippet
    #
    #    def execute(query)
    #      q = newObject 'Query', query
    #      q.exequte
    #    end
    #  end
    #
    #  class MyOleAccessor
    #    it_has_ole_runtime
    #    like_ole_runtime external_connection
    #
    #    include Query
    #
    #    def get_foo
    #      text = 'select * from foo where bar'
    #      result = execute text
    #    end
    #  end
    #
    #  MyOleAccessor.new.get_foo
    def is_ole_snippet
      fail 'Ole snippet must be a Module not a Class' if\
        self.class == Class
      extend AssOle::Snippets::IsSnippet
    end

    # Define class or module which transparent call Ole runtime as self
    # @example
    #   info_base = AssMaintainer::InfoBase.new('', 'File="path"')
    #   class MyOleAccessor
    #     like_ole_runtime thick_app
    #
    #     thick_app.run info_base
    #
    #     def hello(s)
    #       sTring s
    #     end
    #   end
    #
    #   MyOleAccessor.hello('Ass')
    def  like_ole_runtime(runtime)
      it_has_ole_runtime runtime
      case self
      when Class then
        include AssOle::Snippets::LikeOleRuntime
      else
        extend AssOle::Snippets::LikeOleRuntime
      end
    end

    # Define class or module wich include Ole runtime
    # @example
    #   info_base = AssMaintainer::InfoBase.new('', 'File="path"')
    #   class MyOleAccessor
    #     it_has_ole_runtime thick_app
    #
    #     thick_app.run info_base
    #
    #     def hello(s)
    #       ole_connector.sTring s
    #     end
    #   end
    #
    #   MyOleAccessor.hello('Ass')
    def it_has_ole_runtime(runtime)
      case self
      when Class then
        include runtime
      else
        extend runtime
      end
    end

    # Define Ole runtime
    # @example
    #  acct_infobase = AssMaintainer::InfoBase.new('name', 'File="path"')
    #
    #  module AccountingExternal
    #    is_ole_runtime :external
    #  end
    #
    #  AccountingExternal.run acct_infobase
    #
    #  module MyScript
    #    like_ole_runtime acct_infobase
    #
    #    def self.version
    #      Metadata.Version
    #    end
    #  end
    #
    #  puts MyScript.version
    def is_ole_runtime(type)
      fail 'Ole runtime is a Module not a Class' if\
        self.class == Class
      case type
      when :external then
        extend AssOle::Runtimes::App::External
      when :thick
        extend AssOle::Runtimes::App::Thick
      when :thin
        extend AssOle::Runtimes::App::Thin
      when :wp
        extend AssOle::Runtimes::Claster::Wp
      when :agent
        extend AssOle::Runtimes::Claster::Agent
      else
        fail "Invalid runtime #{type}"
      end
    end
  end
end

# Pathch for include DSL
class Module
  include AssOle::DSL
end
