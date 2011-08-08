require "rubygems"
require "rspec"
require "outlook"
require "kconv"

describe Outlook do
	o = Outlook.new
	EntryId = "0000000"
	context "半角コロンが与えられたとき" do
		it "全角コロンが得られる" do
			Kconv.tosjis(o.replace(':')).should ==
				 Kconv.tosjis('：')
		end
	end
	context "半角円マークが与えられたとき" do
		it "全角円マークが得られる" do
			Kconv.tosjis(o.replace('\\')).should ==
				Kconv.tosjis('￥')
		end
	end
	
	context "メール件名が与えられたとき" do
		it ":を：に置き換えた件名が得られる(一括で)" do
			o.replace(Kconv.tosjis('Re:メール読みました:p')).should ==
				 Kconv.tosjis('Re：メール読みました：p')
		end
		it "複数の半角記号を全角記号に置き換えた件名が得られる" do
			o.replace(Kconv.tosjis('Re:正/規/表/現*アスタ?パイプ|わかんない><')).should ==
				 Kconv.tosjis('Re：正／規／表／現＊アスタ？パイプ｜わかんない＞＜')
		end
	end
	
=begin
	context "メール検索をするとき" do
		it "メールを正常に取得できる" do
			o.folders
			entryId = o.getFolderEntryId("12")
			o.searchMail(entryId, "TL")
		end
	end
=end
	
	context "メール全件取得を実行したとき" do
		it "メールを取得する？" do
#			o.mails
		end
	end
	context "メール1件取得を実行したとき" do
		it "メールを取得する？" do
#			o.mail(EntryId)
		end
		it "メールを保存する？" do
#			o.saveMail(o.mail(EntryId))
		end
		it "添付ファイルを保存する？" do
#			o.saveFile(o.mail(EntryId))
		end
	end
	context "ディレクトリを作成するとき" do
		it "作成する？" do
#			o.mkdir(o.mail(EntryId))
		end
	end
end
