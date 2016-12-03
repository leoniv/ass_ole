# encoding: utf-8

module AssOle
  require 'ass_launcher'
  # Helpers for transparency and friendly
  # execute 1C:Enterprise Ole connectors methods as
  # methods a Ryby objects and makes easy for use ruby wrappers over the
  # long and awkward 1C:Enterprise embedded programming language syntax.
  module Snippets
    RUNTIME_CONTEXTS = {
      :IbConnection => :external,
      :WpConnection => :wp,
      :AgentConnection => :agent,
      :ThickApplication => :thick,
      :ThinApplication => :thin,
    }
    class ContextError < StandardError; end
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
          ole_connector.send(method, *args)
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
        fail_if_bad_context obj
        obj.send(:include, MethodMissing) unless obj.include? MethodMissing
        obj.send(:include, Argv) unless obj.include? Argv
        obj.send(:include, WinPath) unless obj.include? WinPath
      end

      def is_helper?(obj)
        obj.class == Module
      end

      def fail_if_bad_context(obj)
        fail ContextError unless context_get.include? runtime_context(obj)
      end

      def runtime_context(obj)
        RUNTIME_CONTEXTS[obj.ole_connector.class.name.split('::').last.to_sym]
      end

      def extended(obj)
        fail_if_bad_context obj
        obj.send(:extend, MethodMissing) unless\
          obj.singleton_class.include? MethodMissing
        obj.send(:extend, Argv) unless\
          obj.singleton_class.include? Argv
        obj.send(:extend, WinPath) unless\
          obj.singleton_class.include? WinPath
      end

      def context(*context)
        fail ContextError if (context - RUNTIME_CONTEXTS.values).size > 0
        @context = context_get + context
      end

      def context_get
        @context ||= []
      end
      private :context_get
    end

    module Examples
      # Snippet for serialize and deserilize 1C objects to xml
      module XMLSerializer
        extend AssOle::Snippets::IsSnippet
        context :external, :thick

        def to_xml(obj)
          zxml = newObject 'XMLWriter'
          zxml.SetString
          xDTOSerializer.WriteXML zxml, obj
          zxml.close
        end

        def to_xml_file(obj, xml_file)
          zxml = newObject 'XMLWriter'
          _path = xml_file.respond_to?(:path) ? xml_file.path : xml_file
          zxml.openFile(win_path(_path))
          xDTOSerializer.WriteXML zxml, obj
          zxml.close
          xml_file
        end

        def from_xml(xml)
          fail 'FIXME'
        end

        def from_xml_file(xml_file)
          fail 'FIXME'
        end
      end

      # Snippet for worcking with 1C Query object
      module Query
        extend AssOle::Snippets::IsSnippet
        context :external, :thick

        def query(text, temp_tables_manager = nil, **params)
          q = newObject('Query', text)
          q.TempTablesManager = temp_tables_manager if\
            temp_tables_manager
          params.each do |k,v|
            q.SetParameter(k.to_s,v)
          end
          q
        end
      end

      module OleHelpers
        module LookingForItems
          include AssOle::Snippets::Examples::Query
          def find_only(query_text, count, **params)
            t = query(query_text, **params).execute.unload
            fail "To many #{self.class.name} found" if\
              t.count > count
            return nil if t.count == 0
            return t
          end
          private :find_only

          def find_only_zero_zero(text, **params)
            t = find_only(text, 1, **params)
            return t.get(0).get(0) if t
          end
          private :find_only_zero_zero

          def fail_not_found(table)
            fail "#{self.class.name} not found" if table.nil?
            table
          end
          private :fail_not_found
        end
      end

      module IsCatalog
        STANDARD_ATTRIBUTES = %w(Ref Code Description Owner Parent IsFolder\
                                DeletionMark Predefined)
        extend AssOle::Snippets::IsSnippet
        context :external, :thick
        include OleHelpers::LookingForItems
        def folder_create_if_not_exists(desc, parent = nil, owner = nil,
                                        **fields)
          res =  folder_get(desc, parent, owner, **fields)
          return res if res
          new_folder(desc, parent, owner, **fields)
        end
        private :folder_create_if_not_exists

        def item_create_if_not_exists(desc, parent = nil, owner = nil,
                                        **fields)
          res =  item_get(desc, parent, owner, **fields)
          return res if res
          new_item(desc, parent, owner, **fields)
        end
        private :item_create_if_not_exists

        def new_folder(desc, parent = nil, owner = nil, **fields)
          new_(:CreateFolder, desc, parent, owner, **fields)
        end

        def new_item(desc, parent = nil, owner = nil, **fields)
          new_(:CreateItem, desc, parent, owner, **fields)
        end

        def new_(method, desc, parent, owner, **fields)
          r = catalogs.send(catalog_name.to_sym).send(method)
          fill(r, desc, parent, owner, **fields)
          r.Write
          r.Ref
        end
        private :new_

        def catalog_name
          fail\
            'Overlord #catalog_name method in Class included IsCatalog snippet'
        end

        def fill(elem, desc, parent, owner, **fields)
          elem.Description = desc
          elem.Owner = owner if owner
          elem.Parent = parent if parent
          fields.each do |f, v|
            elem.send("#{f}=".to_sym, v)
          end
          elem
        end
        private :fill

        def empty_ref
          cAtalogs.send(catalog_name.to_sym).emptyRef
        end

        def item_get(desc, parent = nil, owner = nil, **fields)
          find_only_zero_zero(*select_(desc, parent, owner, false, **fields))
        end
        private :item_get

        def folder_get(desc, parent = nil, owner = nil, **fields)
          find_only_zero_zero(*select_(desc, parent, owner, true, **fields))
        end
        private :folder_get

        def select_(desc, parent = nil, owner = nil,
                       isfolder = nil, **fields)
          _fields = {}
          _fields.merge!(fields)
          _fields[:IsFolder] = isfolder if isfolder
          _fields[:Description] = desc
          _fields[:Parent] = parent if parent
          _fields[:Owner] = owner if owner
          text = "select ref from catalog.#{catalog_name} where\n"
          text << "Description like &description\n"
          text << "and IsFolder = &IsFolder\n" if isfolder
          text << "and Parent = &parent\n" if parent
          text << "and Owner = &owner\n" if owner
          fields.keys.each do |k|
            text << "and #{k} = &#{k}\n"
          end
          [text, _fields]
        end
        private :select_
      end
    end
  end

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
        if self.class == Module
          instance_variable_get(:@ole_connector)
        elsif  self.class == Class
          class_variable_get(:@@ole_connector)
        else
          self.class.ole_connector
        end
      end
    end

    module ModuleMethods
      def ole_connector
        @ole_connector
      end

      def run(connection_string_or_uri)
        ole_connector.__open__ connection_string_or_uri unless runned?
      end

      def stop
        ole_connector.__close__ if runned?
      end

      def runned?
        ole_connector.__opened__?
      end

      def included(obj)
        ole_connector_set obj
        obj.send(:extend, RuntimeDispatcher)
      end

      def extended(obj)
        ole_connector_set obj
      end

      def ole_connector_set(obj)
        if obj.class == Class
          obj.class_variable_set(:@@ole_connector, ole_connector)
        else
          obj.instance_variable_set(:@ole_connector, ole_connector)
        end
      end
      private :ole_connector_set
    end

    # @api private
    module AbstractRuntime
      def platform_require(obj)
        eval "#{obj.name}::PLATFORM_REQUIRE"
      end

      def extended(obj)
        obj.instance_variable_set(:@ole_connector,
                                  ole_class.new(obj.PLATFORM_REQUIRE))
        obj.send(:extend, AssOle::Runtimes::ModuleMethods)
        obj.send(:include, AssOle::Runtimes::RuntimeDispatcher)
        Runtimes.runtimes << obj
      end
    end

    # 1C:Enterprise application runtime helpers
    module App
      # 1C:Enterprise application external connection runtime helper
      module External
        extend AbstractRuntime
        def self.ole_class
          AssLauncher::Enterprise::Ole::IbConnection
        end
      end

      # 1C:Enterprise thick application connection runtime helper
      module Thick
        extend AbstractRuntime
        def self.ole_class
          AssLauncher::Enterprise::Ole::ThickApplication
        end
      end

      # 1C:Enterprise thin application connection runtime helper
      module Thin
        extend AbstractRuntime
        def self.ole_class
          AssLauncher::Enterprise::Ole::ThinApplication
        end
      end
    end

    # 1C:Enterprise server runtime helpers
    module Claster
      # 1C:Enterprise serever worcking process connection helper
      module Wp
        extend AbstractRuntime
        def self.ole_class
          AssLauncher::Enterprise::Ole::WpConnection
        end
      end

      # 1C:Enterprise serever agent connection helper
      module Agent
        extend AbstractRuntime
        def self.ole_class
          AssLauncher::Enterprise::Ole::AgentConnection
        end
      end
    end
  end
end
