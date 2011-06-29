require "rubygems"
require "rspec"
require "msgParse"

describe MsgParse do
	m = MsgParse.new

	context "初期化したとき" do
		it "正常にインスタンスが生成できる" do
			MsgParse.new
		end
	end
	
	context ".msgファイルを読み込んだとき" do
		it "正常に成功する" do
			m.inputMsg("c:/normal.msg")
		end
	end

	context ".msgファイルの添付ファイルを確認し" do
		it "正常にサイズが確認できる" do
			m.inputMsg("c:/notemp.msg")
			m.attachmentSize.should == 0
		end
		it "波ダッシュファイルでもサイズは確認できる" do
			m.inputMsg("c:/waveda.msg")
			m.attachmentSize.should == 1
		end
		it "サイズ0であることが確認できる" do
			m.inputMsg("c:/normal.msg")
			m.attachmentSize.should == 1
		end
	end
	
	context ".msgファイルの添付ファイルをDLするとき" do
		it "正常に成功する" do
			m.inputMsg("c:/normal.msg")
			m.saveFile("c:/")
		end
		it "波ダッシュファイルでも確認できる" do
			m.inputMsg("c:/waveda.msg")
			m.saveFile("c:/")
		end
		it "サイズ0の場合正常に成功する" do
			m.inputMsg("c:/notemp.msg")
			m.saveFile("c:/")
		end
	end
	
end
