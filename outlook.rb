# -*- encoding: UTF-8 -*-
require 'rubygems'
require 'rjb'
require 'msgParse'
require 'win32ole'
require 'date'
require 'kconv'
require 'jcode'
$KCODE="s"

class Outlook
	def initialize
		begin
			# Outlookに接続
			ol = WIN32OLE::connect("Outlook.Application")
		rescue WIN32OLERuntimeError
			putsCurrentMethod("MicrosoftOutlookが起動していません。")
			exit
		else
			desktopJa = Kconv.tosjis("デスクトップ")
			# NameSpace取得(getNameSpaceの引数は"MAPI"のみ)
			@nameSpace = ol.getNameSpace("MAPI")
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
	def folders(count=@defaultCount)
		folders = @nameSpace.Folders
		@folderNum = 1
		folders.each do |f|
			if count < @folderNum then
				break
			end
			GC.start
			putsKconv(f.Name)
			putsKconv("N | CNT | FolderName")
			findFolder(f.EntryId)
		end
	end
	
	# EntryIDを元にメールを取得	
	def mail(entryId)
		@nameSpace.GetItemFromID(entryId)
	end
	
	# EntryIdを元にメールの一覧を取得
	def mails(entryId, count=@defaultCount)
		f = @nameSpace.GetFolderFromID(entryId)
		if f.Items.Count == 0 then
			raise Kconv.tosjis("フォルダにメールがありません。")
		end
		@mailNum = 1
		putsKconv("N | A | SentOn              | Name       | Subject")
		f.Items.each do |mail|
			if count < @mailNum then
				break
			end
			GC.start
			putsKconv("#{@mailNum} | " +
						"#{mail.Attachments.Count} | " +
						"#{mail.SentOn} | " +
						"#{mail.SenderName.unpack('A10')} | " +
						"#{mail.Subject.unpack('A35')}")
			@mail[@mailNum.to_s] = mail.EntryId
			@mailNum += 1
		end
	end

	# フォルダ名を再帰的に取得
	def findFolder(entryId, count=@defaultCount)
		folder(entryId).Folders.each do |f|
			if count < @folderNum then
				break
			end
			begin
				putsKconv("#{@folderNum} | " + 
				      "#{f.Items.Count}通 | " + 
				      "#{f.Name}")
				      # + " | #{f.Parent.Name}"
				@folder[@folderNum.to_s] = f.EntryId
				@folderNum += 1
				findFolder(f.EntryId)
			rescue => ex
				putsCurrentMethod(ex)
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
	
		# 保存フォルダ名を添付ファイル(拡張子無し)に変更
	def renameFolder(mail, fileName)
		receivedTime = Date.strptime(mail.SentOn, "%Y/%m/%d").strftime("%Y%m%d")
		rename = "#{@saveRootPath}" + 
								"#{receivedTime}_#{File.basename(fileName, ".*")}\\"
		if !File.exist?(rename) then
			File.rename(@saveDir, rename)
			putsKconv("フォルダ名変更　:#{rename}")
		end
	end
	
	# @saveRootPath/下にSubject.txtファイルを作成しメール内容を書きこむ
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
			#putsKconv("本文保存　　　　:#{fullPath}")
			putsKconv("本文保存　　　　:#{self.replace(mail.Subject)}")
		end
	end
	
	# 添付ファイル保存
	def saveFile(mail)
		if mail.Attachments.Count != 0 then
			mail.Attachments.each do |item|
				item.SaveAsFile("#{@saveDir}" + 
													"#{self.replace(item.FileName)}")
				putsKconv("添付ファイル保存:#{item.FileName}")
				
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
			putsKconv("添付ファイルはありません。")
		end
	end
	
	# Windowsファイルに使えない記号を変換
	def replace(str)
		str.tr('/:*?"<>|\\', Kconv.tosjis('／：＊？”＜＞｜￥'))
	end
	
	# SJISで出力(コマンドプロンプト用)
	def putsKconv(str)
		puts Kconv.tosjis(str)
	end
	
	# エラーの起きたメソッドを出力する
	# 参考http://www.rubyist.net/~nobu/t/20051013.html#p02
	def putsCurrentMethod(ex="")
		puts Kconv.tosjis("Error " + caller.first[/:in \`(.*?)\'\z/, 1] + " - " + ex)
		#puts "Error " + caller.first[/:in \`(.*?)\'\z/, 1] + " - " + ex
	end
end
