# -*- encoding: UTF-8 -*-
require 'win32ole'
require 'date'
require 'jcode'
require 'pp'
$KCODE="s"

class Outlook
	def initialize
		begin
			# Outlookに接続
			ol = WIN32OLE::connect("Outlook.Application")
		rescue WIN32OLERuntimeError
			puts "MicrosoftOutlookが起動していません。"
			exit
		else
			# NameSpace取得(getNameSpaceの引数は"MAPI"のみ)
			@nameSpace = ol.getNameSpace("MAPI")
			# 保存パス指定
			@saveRootPath = "#{ENV["USERPROFILE"]}\\デスクトップ\\"
			# 保存パスに作成するディレクトリ作成
			@saveDir = ""
			# フォルダ選択番号
			@folderNum = -1
			# 番号・フォルダEntryId対応ハッシュ
			@folder = Hash.new
			# メール選択番号
			@mailNum = -1
			# 番号・メールEntryId対応ハッシュ
			@mail = Hash.new
		end
	end
	
	# ハッシュを元にフォルダのEntryIdを取得
	def getFolderEntryId(key)
		@folder[key]
	end
	
	# ハッシュを元にメールのEntryIdを取得
	def getMailEntryId(key)
		@mail[key]
	end

	# EntryIdを元にフォルダを取得
	def folder(entryId)
		@nameSpace.GetFolderFromID(entryId)
	end
	
	# ルートからフォルダの一覧を取得
	def folders
		folders = @nameSpace.Folders
		@folderNum = 1
		folders.each do |f|
			GC.start
			puts f.Name
			findFolder(f.EntryId)
		end
	end
	
	# EntryIDを元にメールを取得	
	def mail(entryId)
		@nameSpace.GetItemFromID(entryId)
	end
	
	# EntryIdを元にメールの一覧を取得
	def mails(entryId)
		f = @nameSpace.GetFolderFromID(entryId)
		if f.Items.Count == 0 then
			raise "フォルダにメールがありません。"
		end
		@mailNum = 1
		puts "N | A | SentOn | Name | Subject"
		f.Items.each do |mail|
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

	# フォルダ名を再帰的に取得
	def findFolder(entryId)
		folder(entryId).Folders.each do |f|
			begin
				puts "#{@folderNum} #{f.Name}"
				@folder[@folderNum.to_s] = f.EntryId
				@folderNum += 1
				# puts f.Parent.Name
				findFolder(f.EntryId)
			rescue => ex
				puts ex
			end
		end
	end
	
	# @saveRootPath下にディレクトリ作成
	def mkdir(mail)
		# 受信日のYYYYMMDD
		receivedTime = Date.strptime(mail.SentOn, "%Y/%m/%d").strftime("%Y%m%d")
		@saveDir = "#{@saveRootPath}" +
				"#{receivedTime}_#{replace(mail.Subject)}\\"
		if !File.exist?(@saveDir) then
			Dir.mkdir(@saveDir)
		end
	end
	
	# @saveRootPath/下にSubject.txtファイルを作成しメール内容を書きこむ
	def saveMail(mail)
		fullPath = "#{@saveDir}#{self.replace(mail.Subject)}.txt"
		File.open(fullPath, "w") do |file|
			file.write "SENDER : #{mail.SenderName}" +
					"(#{mail.SenderEmailAddress})\n"
			file.write "TO : #{mail.To}\n"
			file.write "CC : #{mail.CC}\n"
			file.write "ReceivedTime: #{mail.SentOn}\n"
			file.write "SUBJECT : #{mail.Subject}\n"
			file.write "BODY : \n#{mail.Body}\n"
			puts "本文を保存しました。(#{fullPath})"
		end
	end

	# 添付ファイル保存
	def saveFile(mail)
		if mail.Attachments.Count != 0 then
			mail.Attachments.each do |item|
				item.SaveAsFile("#{@saveDir}" +
					"#{self.replace(item.FileName)}")
				puts "#{item.FileName}を保存しました。"
			end
		else
			puts "添付ファイルはありません。"
		end
	end

	# Windowsファイルに使えない記号を変換
	def replace(str)
		str.tr('/:*?"<>|\\', '／：＊？”＜＞｜￥')
	end
end
