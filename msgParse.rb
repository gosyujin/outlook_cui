# -*- encoding: UTF-8 -*-
require 'rubygems'
require 'rjb'
require 'kconv'

class MsgParse
	include Rjb
	def initialize
		#Rjb::load('./')
		initJavaClass
		addJar
		@msg = nil
	end
	
	# JavaのクラスをImport
	def initJavaClass
		@system = import("java.lang.System")
		@string = import("java.lang.String")
		@list = import("java.util.List")
		@fileOutputStream = import("java.io.FileOutputStream")
	end
	
	# jarをImport
	def addJar
		Rjb::add_jar(File.expand_path('tnef-1.3.1.jar'))
		Rjb::add_jar(File.expand_path('poi-3.2-FINAL-20081019.jar'))
		Rjb::add_jar(File.expand_path('msgparser-1.10.jar'))
		@msgParser = import("com.auxilii.msgparser.MsgParser")
		@fileAttachment = import("com.auxilii.msgparser.attachment.FileAttachment")
	end
	
	# .msgファイル読み込み
	def inputMsg(path)
		@msg = @msgParser.new.parseMsg(path)
	end
	
	# 添付ファイルの数をカウント
	def attachmentSize
		@msg.getAttachments.size
	end
	
	# 添付ファイルをpathに保存
	# 返り値に保存した添付ファイル名(の一つ)
	def saveFile(path)
		fileName = ""
		if attachmentSize != 0 then
			for i in 0..attachmentSize - 1
				file = @msg.getAttachments.get(i)
				begin
					fileName = file.getLongFilename
				rescue => ex
					puts "File name is includeing WAVE DASH??:#{ex}"
					fileName = file.getFilename
				end
				out = @fileOutputStream.new(path + fileName)
				out.write(file.getData)
				puts "Complete        :#{fileName}"
				out.close
			end
		else
			puts "no temp file."
			return nil
		end
		return fileName
	end
end
