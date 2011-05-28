require 'outlook'

def getMails(o)
	# フォルダ一覧表示
	o.folders
	begin
		key = ""
		# Input 1〜...
		# 文字列のto_iは0となる
		while key.to_i < 1
			print "Input Folder No.:"
			key = STDIN.gets.chomp
		end
		# 先頭の0取り
		key = key.to_i.to_s
		entryId = o.getFolderEntryId(key)
		# メール一覧表示
		o.mails(entryId)
	rescue => ex
		puts ex
		# フォルダ取得ミスったらやり直し
		retry
	end
end

def getMail(o)
	#TODO 複数できるように
	# Input 1〜...
	# 文字列のto_iは0となる
	key = ""
	while key.to_i < 1
		print "Input Mail No.:"
		key = STDIN.gets.chomp
	end

	# 先頭の0取り
	key = key.to_i.to_s
	entryId = o.getMailEntryId(key)
	# 対象メール取得
	mail = o.mail(entryId)

	# 保存
	o.mkdir(mail)
	o.saveMail(mail)
	o.saveFile(mail)
end

o = Outlook.new
getMails(o)
getMail(o)