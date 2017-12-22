require 'test_helper'

module AssOleTest
  describe StandardError do
    HRESULT_ASCII_MESSAGE =
      "unknown property or method: `Write'\n"\
      "Текст по русски"\
      "HRESULT error code:0x80020006\n"\
      "      \xCD\xE5\xE8\xE7\xE2\xE5\xF1\xF2\xED\xEE\xE5 \xE8\xEC\xFF.\n"
        .force_encoding('ASCII-8BIT')

    patched_error = Class.new StandardError do
      def initialize
        super hresult_mess
      end

      def hresult_mess
        HRESULT_ASCII_MESSAGE
      end

      __patch_message__
    end

    not_patched_error = Class.new StandardError do
      def initialize
        super hresult_mess
      end

      def hresult_mess
        HRESULT_ASCII_MESSAGE
      end
    end

    it 'Fail conact string for unpatched Error' do
      e = proc {
        [not_patched_error.new.message, "Текст по русски"].join("\n")
      }.must_raise Encoding::CompatibilityError
      e.message.must_match\
        %r{incompatible character encodings: ASCII-8BIT and UTF-8}i
    end

    it 'Not fail conact string for patched Error' do
      [patched_error.new.message, "Текст по русски"].join("\n")
    end

    it '#message' do
      inst = patched_error.new
      inst.message.must_equal "unknown property or method: `Write'\n"\
                              "Текст по русски"\
                              "HRESULT error code:0x80020006"\
    end

    it 'NoMethodError message patched' do
      inst = NoMethodError.new HRESULT_ASCII_MESSAGE
      inst.message.must_equal "unknown property or method: `Write'\n"\
                              "Текст по русски"\
                              "HRESULT error code:0x80020006"\
    end

    it 'WIN32OLERuntimeError message patched' do
      inst = WIN32OLERuntimeError.new HRESULT_ASCII_MESSAGE
      inst.message.must_equal "unknown property or method: `Write'\n"\
                              "Текст по русски"\
                              "HRESULT error code:0x80020006"\
    end
  end
end
