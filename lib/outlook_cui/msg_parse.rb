# -*- encoding: UTF-8 -*-
require "rjb"

module OutlookCui
  class MsgParse
    include Rjb

    Jar_Path = "lib/vendor"
    Jar_Files = [ "tnef-1.3.1.jar", 
                  "poi-3.2-FINAL-20081019.jar", 
                  "msgparser-1.10.jar" ]

    def initialize
      Rjb::load(classpath = '.', jvmargs=[])
      import_java_package
      add_jar
    end

    def down(msg_path)
      msg = @msg_parser.new.parseMsg(msg_path.tosjis)
      attach_count = msg.getAttachments.size

      attach_count.times do |i|
        file = msg.getAttachments.get(i)
        filename = self.replace(file.getLongFilename)
        path = self.pathname(File::dirname(msg_path.tosjis), filename.tosjis)

        out = @file_output_stream.new(path)
        out.write(file.getData)
        out.close
	puts ".msg unzip : -> #{filename}"
      end
    end

private
    def import_java_package
      @system             = import('java.lang.System')
      @string             = import("java.lang.String")
      @list               = import("java.util.List")
      @file_output_stream = import("java.io.FileOutputStream")
    end

    def add_jar
      Jar_Files.each do |jar|
        Rjb::add_jar(File.expand_path(self.pathname(Jar_Path, jar)))
      end
      @msg_parser      = import('com.auxilii.msgparser.MsgParser')
      @file_attachment = import('com.auxilii.msgparser.attachment.FileAttachment')
    end
  end
end

class String
  def tosjis
    self.encode("Shift_JIS")
  end
end
