# -*- encoding: UTF-8 -*-
require 'outlook'

def inputCheck(message)
	# Input 1～...
	# 文字列のto_iは0となる
	print message
	key = STDIN.gets.chomp
	# 先頭の0取り
	key = key.to_i.to_s
end

def getMails(o)
	# フォルダ一覧表示
	o.folders
	begin
		key = ""
		while key.to_i < 1
		
			key = inputCheck("Input Folder No.:")
		end
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
	# Input 1～...
	# 文字列のto_iは0となる
	keys = Array.new
	begin	
		print "Input Mail No.(split space):"
		input = STDIN.gets.chomp
		keys.clear
		# 先頭の0取り
		input.split(" ").each do |k|
			keys << k.to_i.to_s
		end
		beforeKeysLength = keys.length
		keys.reject! do |x|
			x == "0"
		end
	end while beforeKeysLength != keys.length
	
	keys.each do |key|
		entryId = o.getMailEntryId(key)
		
		# 対象メール取得
		mail = o.mail(entryId)
		
		#	mail.ItemProperties.each do |e|
		#		p e.Name
		#	end
		
		# 保存
		o.mkdir(mail)
		o.saveMail(mail)
		o.saveFile(mail)
	end
end

o = Outlook.new
while 1 
	getMails(o)
	while 1
		getMail(o)
	end
end