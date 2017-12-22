# Monkey patch for encoding error message
class StandardError
  def self.__patch_message__
    old_message = instance_method(:message)
    define_method(:message) do
      old_message.bind(self).call.force_encoding('ASCII-8BIT')\
        .split(%r{(HRESULT error code:\dx\d+)}i)[0 .. 1].join
        .force_encoding('UTF-8').strip
    end
  end
end
