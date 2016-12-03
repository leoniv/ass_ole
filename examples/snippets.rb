module AssOleExamples
  require 'ass_ole'
  module Snippets
    # Snippet for serialize and deserilize 1C objects to xml
    module XMLSerializer
      extend AssOle::Snippets::IsSnippet

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
