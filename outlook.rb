# -*- encoding: UTF-8 -*-
# = MicrosoftOutlook内のメールを操作するクラス
require 'rubygems'
require 'rjb'
require 'msgParse'
require 'win32ole'
require 'date'
require 'kconv'
require 'jcode'
$KCODE="s"

class Outlook
	# MicrosoftOutlookに接続後初期化処理を行う。
	# MicrosoftOutlookが起動していないと終了する。
	def initialize
		begin
			@ol = WIN32OLE::connect("Outlook.Application")
		rescue WIN32OLERuntimeError
			putsError("MicrosoftOutlookが起動していません。")
			exit
		else
			desktopJa = Kconv.tosjis("デスクトップ")
			# NameSpace取得(getNameSpaceの引数は"MAPI"のみ)
			@nameSpace = @ol.getNameSpace("MAPI")
			# 保存パス指定
			@saveRootPath = "#{ENV["USERPROFILE"]}\\" + desktopJa + "\\"
			# 保存パスに作成するディレクトリ作成
			@saveDir = ""
			# フォルダ選択番号、ハッシュ
			@folderNum = -1
			@folder = Hash.new
			# メール選択番号、ハッシュ
			@mailNum = -1
			@mail = Hash.new
			# 取得件数のデフォルト値
			@defaultCount = 20
		end
	end
	
	# プラットフォームがWindowsの場合標準出力をKconv.tosjisでラップ
	if RUBY_PLATFORM.downcase =~ /mswin(?!ce)|mingw|cygwin|bccwin/ then
		def $stdout.write(str)
			super Kconv.tosjis(str)
		end
	end
	
	# keyを元にフォルダのEntryIdを取得する
	def getFolderEntryId(key)
		@folder[key]
	end
	
	# keyを元にメールのEntryIdを取得する
	def getMailEntryId(key)
		@mail[key]
	end

	# entryIdを元にフォルダを取得する
	def folder(entryId)
		@nameSpace.GetFolderFromID(entryId)
	end
	
	# ルートからフォルダの一覧を取得する
	def folders(count=@defaultCount)
		folders = @nameSpace.Folders
		@folderNum = 1
		folders.each do |f|
			if count < @folderNum then
				break
			end
			GC.start
			puts f.Name
			puts "N | CNT | FolderName"
			findFolder(f.EntryId)
		end
	end
	
	# entryIDを元にメールを取得する
	def mail(entryId)
		@nameSpace.GetItemFromID(entryId)
	end
	
	# entryIdを元に対象フォルダのメール一覧を取得する
	def mails(entryId, isAttachmentOnlyMode, count=@defaultCount)
		f = @nameSpace.GetFolderFromID(entryId)
		if f.Items.Count == 0 then
			raise "フォルダにメールがありません。"
		end
		@mailNum = 1
	    puts "N | A | SentOn              | Name       | Subject"
		f.Items.each do |mail|
			if count < @mailNum then
				break
			end
			if isAttachmentOnlyMode && mail.Attachments.Count == 0 then 
				next
			end
			GC.start
			puts "#{@mailNum} | " +
						"#{mail.Attachments.Count} | " +
						"#{mail.SentOn} | " +
						"#{mail.SenderName.unpack('A10')} | " +
						"#{mail.Subject.unpack('A35')}"
			@mail[@mailNum.to_s] = mail.EntryId
			@mailNum += 1
		end
	end

	# entoryIdを元にフォルダ名を再帰的に取得する
	def findFolder(entryId, count=@defaultCount)
		folder(entryId).Folders.each do |f|
			if count < @folderNum then
				break
			end
			begin
				puts "#{@folderNum} | " + 
					  "#{f.Items.Count}通 | " + 
					  "#{f.Name}"
					  # + " | #{f.Parent.Name}"
				@folder[@folderNum.to_s] = f.EntryId
				@folderNum += 1
				findFolder(f.EntryId)
			rescue => ex
				putsError(ex)
			end
		end
	end
	
	def searchMail(entryId, subject)
		folder = "ML"
		# WIN32OLEインスタンス, イベントインタフェース名
		events = WIN32OLE_EVENT.new(@ol, "ApplicationEvents_11")
		@ol.AdvancedSearch(folder, "urn:schemas:mailheader:subject LIKE '#{subject}'")
		
	end
	
	# @saveRootPath下にディレクトリを作成する
	def mkdir(mail)
		# 受信日のYYYYMMDD
		receivedTime = DateTime.strptime(mail.SentOn, "%Y/%m/%d %H:%M:%S").strftime("%Y%m%d_%H%M%S")
		@saveDir = "#{@saveRootPath}" + 
								"#{receivedTime}_#{replace(mail.Subject)}\\"
		if !File.exist?(@saveDir) then
			Dir.mkdir(@saveDir)
		end
	end
	
	# 保存フォルダ名を添付ファイル(拡張子無し)に変更する
	def renameFolder(mail, fileName)
		receivedTime = DateTime.strptime(mail.SentOn, "%Y/%m/%d %H:%M:%S").strftime("%Y%m%d_%H%M%S")
		rename = "#{@saveRootPath}" + 
								"#{receivedTime}_#{File.basename(fileName, ".*")}\\"
		if !File.exist?(rename) then
			File.rename(@saveDir, rename)
			puts "■フォルダ名変更　:#{rename}"
		end
	end
	
	# @saveDir下にSubject.txtファイルを作成しメール内容を書きこむ
	def saveMail(mail)
		fullPath = "#{@saveDir}#{self.replace(mail.Subject)}.txt"
		File.open(fullPath, "w") do |file|
			file.write "SENDER      : #{mail.SenderName}" + 
									"(#{mail.SenderEmailAddress})\n"
			file.write "TO          : #{mail.To}\n"
			file.write "CC          : #{mail.CC}\n"
			file.write "ReceivedTime: #{mail.SentOn}\n"
			file.write "SUBJECT     : #{mail.Subject}\n"
			file.write "BODY        : \n#{mail.Body}\n"
			#puts "■本文保存　　　　:#{fullPath}"
			puts "■本文保存　　　　:#{self.replace(mail.Subject)}"
		end
	end
	
	# @saveDir下に添付ファイルを保存する
	def saveFile(mail)
		if mail.Attachments.Count != 0 then
			mail.Attachments.each do |item|
				item.SaveAsFile("#{@saveDir}#{self.replace(item.FileName)}")
				puts "■添付ファイル保存:#{item.FileName}"
				
				# ファイルが.msgだった場合添付ファイルをぶっこぬき
				# フォルダ名も変更
				if item.FileName =~ /.*\.msg/ then 
					msg = MsgParse.new
					msg.inputMsg("#{@saveDir}#{self.replace(item.FileName)}")
					fileName = msg.saveFile(@saveDir)
					renameFolder(mail, fileName)
				end
			end
		else
			puts "添付ファイルはありません。"
		end
	end
	
	# Windowsファイルに使えない記号を変全角に変換する
	def replace(str)
		str.tr('/:*?"<>|\\', Kconv.tosjis('／：＊？”＜＞｜￥'))
	end
	
	# エラーの起きたメソッドを出力する
	# http://www.rubyist.net/~nobu/t/20051013.html#p02
	def putsError(ex="")
		puts "[Error]" + caller.first[/:in \`(.*?)\'\z/, 1] + " - " + ex
	end
end
