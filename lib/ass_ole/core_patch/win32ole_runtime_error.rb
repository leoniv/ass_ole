require 'ass_ole/core_patch/standard_error'
# Monkey patch for encoding error message
class WIN32OLERuntimeError
  __patch_message__
end
