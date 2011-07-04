require "rubygems"
require "rspec"
require "outlook"

describe Outlook do
	o = Outlook.new
	EntryId = "0000000"
	context "半角コロンが与えられたとき" do
		it "全角コロンが得られる" do
			o.replace(':').should ==
				 '：'
		end
	end
	context "半角円マークが与えられたとき" do
		it "全角円マークが得られる" do
			o.replace('\\').should ==
				'￥'
		end
	end
	
	context "メール件名が与えられたとき" do
		it ":を：に置き換えた件名が得られる(一括で)" do
			o.replace('Re:メール読みました:p').should ==
				 'Re：メール読みました：p'
		end
		it "複数の半角記号を全角記号に置き換えた件名が得られる" do
			o.replace('Re:正/規/表/現*アスタ?パイプ|わかんない><').should ==
				 'Re：正／規／表／現＊アスタ？パイプ｜わかんない＞＜'
		end
	end
	
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