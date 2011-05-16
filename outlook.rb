# -*- encoding: UTF-8 -*-
require "win32ole"

# 保存ディレクトリ基準。一応マイドキュメントへ
SaveRootPath = "#{ENV["USERPROFILE"]}\\My Documents"

class Outlook
	def initialize
		# Outlookに接続
		ol = WIN32OLE::connect("Outlook.Application")
		# NameSpace取得(getNameSpaceの引数は"MAPI"のみ)
		@nameSpace = ol.getNameSpace("MAPI")
	end

	# EntryIDを元にメールを取得	
	def mail(entryId)
		item = @nameSpace.GetItemFromID(entryId)
		return item
	end
	
	# メールの一覧を取得
	def mails
		# GetDefaultFolder(6)は受信トレイ
		folder = @nameSpace.GetDefaultFolder(6)
		folder.Items.each do |mail|
			GC.start
			yield mail
		end
	end
		
end

outlook = Outlook.new
# メールのタイトル、送信日時、EntryID一欄を表示
outlook.mails do |mail|
	#puts "To     :#{mail.To}"
	puts "#{mail.SentOn} | #{mail.Subject.unpack("a50")}"
	puts "#{mail.EntryID}"
end

# 入力待ち。保存したいメールのEntryIdを貼りつけ
mail = outlook.mail(STDIN.gets.chomp)
#puts mail.Attachments.ole_methods
#puts mail.Attachments.Item(1).ole_methods

# 添付ファイルのファイル名を取得
file = mail.Attachments.Item(1).FileName
# RootPath下に添付ファイル名のディレクトリ作成
SaveDir = "#{SaveRootPath}\\#{file.split(".")[0]}"
Dir.mkdir(SaveDir)

# 添付ファイル保存
mail.Attachments.Item(1).SaveAsFile("#{SaveDir}\\#{file}")

# ディレクトリ下にタイトル.txtファイルを作成し送信者、タイトル、本文を書きこむ
File.open("#{SaveDir}\\#{mail.Subject}.txt", "w") do |file|
	file.write "SENDER      : #{mail.SenderName}(#{mail.SenderEmailAddress})\n"
	file.write "TO          : #{mail.To}\n"
	file.write "CC          : #{mail.CC}\n"
	file.write "ReceivedTime: #{mail.SentOn}\n"
	file.write "SUBJECT     : #{mail.Subject}\n"
	file.write "BODY: \n#{mail.Body}\n"
end
