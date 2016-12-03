module AssOle
  # Helpers for transparency and friendly
  # execute 1C:Enterprise Ole connectors methods as
  # methods a Ryby objects and makes easy for use ruby wrappers over the
  # long and awkward 1C:Enterprise embedded programming language syntax.
  # @api private
  module Snippets
    GOOD_CONTEXT = AssLauncher::Enterprise::Ole::OLE_CLIENT_TYPES.values

    class ContextError < StandardError
      def initialize(ole_class)
        super "Invalid `ole_connector': #{ole_class}"
      end
    end

    def self.fail_if_bad_context(obj)
      fail ContextError.new(ole_class(obj)) unless good_context? obj
    end

    def self.good_context?(obj)
      GOOD_CONTEXT.include? ole_class(obj)
    end
    private_class_method :good_context?

    def self.ole_class(obj)
      obj.ole_connector.class
    end
    private_class_method :ole_class

    module IsSnippet
      # Helper for pass into 1C ole runtime understandable file system pathes.
      # This provides method `win_path` which will be available in snippets
      module WinPath
        def win_path(path)
          AssLauncher::Support::Platforms.path(path).win_string
        end
      end

      # It worcking via method_missing handler
      module MethodMissing
        def method_missing(method, *args)
          AssOle::Snippets.fail_if_bad_context(self)
          return ole_connector.send(method, *args) if ole_connector
        end
      end

      # Ole ARGV helper for get value over parameters
      # This provides method `argv` which will be available in snippets
      module Argv
        def argv(i)
          WIN32OLE::ARGV[i]
        end
      end

      def included(obj)
        return if is_helper?(obj)
        obj.send(:include, MethodMissing) unless obj.include? MethodMissing
        obj.send(:include, Argv) unless obj.include? Argv
        obj.send(:include, WinPath) unless obj.include? WinPath
      end

      def is_helper?(obj)
        obj.class == Module
      end

      def extended(obj)
        obj.send(:extend, MethodMissing) unless\
          obj.singleton_class.include? MethodMissing
        obj.send(:extend, Argv) unless\
          obj.singleton_class.include? Argv
        obj.send(:extend, WinPath) unless\
          obj.singleton_class.include? WinPath
      end

      extend self
    end
  end
end
