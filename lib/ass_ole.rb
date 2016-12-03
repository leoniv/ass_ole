# encoding: utf-8

module AssOle
  require 'ass_launcher'
  require 'ass_ole/version'
  require 'ass_ole/snippets'
  require 'ass_ole/dsl'

  # Helpers for friendly uses 1C:Enterprise Ole connectors
  # @example
  #  module InfoBase1ExternalRuntime
  #    def self.PLATFORM_REQUIRE
  #      '~> 8.3.8'
  #    end
  #    extend AssOle::Runtimes::App::External
  #  end
  #
  #  class MyClassWitRuntime
  #    def ass_application_name
  #      ole_connector.methadata.name
  #    end
  #  end
  #  connection_string = 'File="./tmp/info_base1.ib"'
  #  InfoBase1ExternalRuntime.run(connection_string)
  #  puts MyClassWitRuntime.new.ass_application_name
  module Runtimes
    # @api private
    def self.runtimes
      @runtimes ||= []
    end

    # @api private
    # `at_exit` handler closed all opened connections
    def self.do_at_exit
      runtimes.each do |r|
        r.stop
      end
    end

    at_exit do
      do_at_exit
    end

    module RuntimeDispatcher
      def ole_connector
        ole_runtime_get.ole_connector
      end

      def ole_runtime_get
        if self.class == Module
          instance_variable_get(:@ole_runtime)
        elsif  self.class == Class
          class_variable_get(:@@ole_runtime)
        else
          self.class
        end
      end
      private :ole_runtime_get
    end

    module ModuleMethods
      attr_reader :ole_connector

       def run_(connection_string_or_uri)
         ole_connector.__open__ connection_string_or_uri unless runned?
       end
       private :run_

      def stop
        ole_connector.__close__ if runned?
      end

      def runned?
        initialized? && ole_connector.__opened__?
      end

      def initialized?
        !ole_connector.nil?
      end

      def included(obj)
        ole_runtime_set obj
        obj.send(:extend, RuntimeDispatcher)
      end

      def extended(obj)
        ole_runtime_set obj
      end

      def ole_runtime_set(obj)
        if obj.class == Class
          obj.class_variable_set(:@@ole_runtime, self)
        else
          obj.instance_variable_set(:@ole_runtime, self)
        end
      end
      private :ole_runtime_set
    end

    # @api private
    module AbstractRuntime
      def extended(obj)
        obj.send(:extend, AssOle::Runtimes::ModuleMethods)
        obj.send(:include, AssOle::Runtimes::RuntimeDispatcher)
        Runtimes.runtimes << obj
      end
    end

    # 1C:Enterprise application runtime helpers
    module App
      module Abstract
        def run(info_base)
          return ole_connector if runned?
          instance_variable_set(:@ole_connector, info_base.ole(ole_type))
          run_ info_base.connection_string
        end
      end
      # 1C:Enterprise application external connection runtime helper
      module External
        extend AbstractRuntime
        include Abstract
        def ole_type
          :external
        end
      end

      # 1C:Enterprise thick application connection runtime helper
      module Thick
        extend AbstractRuntime
        include Abstract
        def ole_type
          :thick
        end
      end

      # 1C:Enterprise thin application connection runtime helper
      module Thin
        extend AbstractRuntime
        include Abstract
        def ole_type
          :thin
        end
      end
    end

    # 1C:Enterprise server runtime helpers
    module Claster
      module Abstract
        def run(uri, platform_require = '> 0')
          return ole_connector if runned?
          instance_variable_set(:@ole_connector,ole_class.new(platform_require))
          run_ uri
        end
      end
      # 1C:Enterprise serever worcking process connection helper
      module Wp
        extend AbstractRuntime
        include Abstract
        def ole_class
          AssLauncher::Enterprise::Ole::WpConnection
        end
      end

      # 1C:Enterprise serever agent connection helper
      module Agent
        extend AbstractRuntime
        include Abstract
        def ole_class
          AssLauncher::Enterprise::Ole::AgentConnection
        end
      end
    end
  end
end
