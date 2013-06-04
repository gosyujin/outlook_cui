# -*- encoding: utf-8 -*-

module Outlook
  module Utility
    extend self

    def dummy
      puts "dummy"
    end

    # my rjust
    def rjust(num, length=nil)
      # fix me
      length ||= 3
      num.to_s.rjust(length)
    end

    # use Pathname return path to string
    def pathname(parent, child)
      (Pathname(parent) + child).to_s
    end

    # replace "_" if not used Window sign
    def replace(str)
      str.tr(' /	　:*?"<>|\\', '_')
    end
  end
end
