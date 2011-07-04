# -*- encoding: UTF-8 -*-
# = .msgファイルを操作するクラス
require 'rubygems'
require 'rjb'
require 'kconv'

class MsgParse
	include Rjb
	
	# JavaクラスのImport、Jarの読み込み
	# 初期化処理を行う
	def initialize
		#Rjb::load('./')
		initJavaClass
		addJar
		@msg = nil
	end
	
	# プラットフォームがWindowsの場合標準出力をKconv.tosjisでラップ
	if RUBY_PLATFORM.downcase =~ /mswin(?!ce)|mingw|cygwin|bccwin/ then
		def $stdout.write(str)
			super Kconv.tosjis(str)
		end
	end
	
	# JavaのクラスをImportする
	def initJavaClass
		@system = import("java.lang.System")
		@string = import("java.lang.String")
		@list = import("java.util.List")
		@fileOutputStream = import("java.io.FileOutputStream")
	end
	
	# Jarを読みこむ
	def addJar
		Rjb::add_jar(File.expand_path('lib/tnef-1.3.1.jar'))
		Rjb::add_jar(File.expand_path('lib/poi-3.2-FINAL-20081019.jar'))
		Rjb::add_jar(File.expand_path('lib/msgparser-1.10.jar'))
		@msgParser = import("com.auxilii.msgparser.MsgParser")
		@fileAttachment = import("com.auxilii.msgparser.attachment.FileAttachment")
	end
	
	# .msgファイルを読みこむ
	def inputMsg(path)
		@msg = @msgParser.new.parseMsg(path)
	end
	
	# .msgファイルの添付ファイル数をカウントする
	def getAttachmentSize
		@msg.getAttachments.size
	end
	
	# 添付ファイルをpathに保存する
	# 返り値は保存した添付ファイル名(の一つ)
	def saveFile(path)
		fileName = ""
		if getAttachmentSize != 0 then
			for i in 0..getAttachmentSize - 1
				file = @msg.getAttachments.get(i)
				begin
					fileName = file.getLongFilename
				rescue => ex
					puts "File name is including WAVE DASH?:#{ex}"
					fileName = file.getFilename
				end
				out = @fileOutputStream.new(path + fileName)
				out.write(file.getData)
				puts "■.msgファイル抽出:#{fileName}"
				out.close
			end
		else
			puts "no temp file."
		end
		return fileName
	end
end
