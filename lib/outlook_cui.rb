# -*- encoding: utf-8 -*-
require "outlook_cui/utility"
require "outlook_cui/version"

require "win32ole"
require "date"

module OutlookCui
  extend self

  Save_Dir_Root = "./mail"
  OS_Limit_Path_Length = 255

  @attachment_only = ENV['attach'].nil? ? false : true
  @limit_count     = ENV['limit'].nil?  ? nil   : ENV['limit'].to_i
  @limit           = @limit_count.nil?  ? false : true
  puts "Attachment only mode"            if @attachment_only
  puts "Limit max #{@limit_count} mails" if @limit

  @f_index = 0
  @folder_list = {}
  @mail_list = {}

  begin
    outlook = WIN32OLE::connect("Outlook.Application")
    @namespace = outlook.getNameSpace("MAPI")
    @folders = @namespace.Folders
  rescue WIN32OLERuntimeError => ex
    puts "Error: Please start Outlook."
    puts "exit."
    puts ex
    exit 1
  end

  def folders(folders=nil)
    recur_folders(folders)
    print "\n"
    @folder_list
  end

  def mails(entry_id)
    folder = folder(entry_id)

    index = 0
    max = folder.Items.count

    folder.Items.each do |mail|
      break if @limit and index == @limit_count

      attach_count = mail.Attachments.Count
      next  if @attachment_only and attach_count == 0

      index += 1
      sent         = ole_call(mail, "SentOn")
      sender       = ole_call(mail, "SenderEmailAddress")
      subject      = ole_call(mail, "Subject")
      entry_id     = ole_call(mail, "EntryId")

      mail = { "attach_count" => attach_count, 
               "sent"         => sent, 
               "sender"       => sender, 
               "subject"      => subject, 
               "entry_id"     => entry_id }
      @mail_list[index.to_s] = mail
      GC.start
      print "\r#{self.rjust(index.to_s)}/#{max} mails read..."
    end
    print "\n"
    @mail_list
  rescue => ex
    puts "Error: mails"
    puts ex
    exit 1
  end

  def save_mail(entry_id, save_dir_root, attachment=true)
    mail = mail(entry_id)
    sender_name   = ole_call(mail, "SenderName")
    sender_email  = ole_call(mail, "SenderEmailAddress")
    to            = ole_call(mail, "To")
    cc            = ole_call(mail, "CC")
    received_time = ole_call(mail, "SentOn")
    subject       = ole_call(mail, "Subject")
    body          = ole_call(mail, "Body")

    save_dir_root ||= Save_Dir_Root
    save_dir_root = File.expand_path(save_dir_root.encode("utf-8"))

    # puts "save_dir   : #{save_dir_root}"
    save_dir_name = "#{received_time.strftime("%Y%m%d_%H%M%S")}_#{self.replace(subject)}"
    save_dir = self.pathname(save_dir_root, save_dir_name)

    save_file_name = "#{save_dir_name}.txt"
    save_file = self.pathname(save_dir, save_file_name)

    if FileTest.exist?(save_dir)
      # delete this directory if you want redownload
      puts "skip(EXIST): #{save_dir_name} is EXIST"
      return
    elsif save_file.length > OS_Limit_Path_Length
      puts "skip(PATH) : TOO LONG length directory path (more than #{OS_Limit_Path_Length})"
      puts " path is   : -> #{save_file}"
      return
    else
      # sleep when downloaded
      sleep(1)
    end

    FileUtils.mkdir_p(save_dir) 
    File.open(save_file, "w") do |file|
      file.write "SENDER      : #{sender_name}\n" \
                 "              #{sender_email}\n" \
                 "TO          : #{to}\n" \
                 "CC          : #{cc}\n" \
                 "ReceivedTime: #{received_time}\n" \
                 "SUBJECT     : #{subject}\n" \
                 "BODY        : \n" \
                 "#{body}\n"
    end
    puts "save_mail  : #{subject}"

    if attachment then
      save_attach(entry_id, save_dir)
    end
    GC.start
  rescue => ex
    puts "Error: save_mail"
    puts ex
    exit 1
  end

  def save_attach(entry_id, save_dir)
    # puts "save_dir   : #{save_dir}"
    mail = mail(entry_id)
    attach_count = mail.Attachments.Count 
    attachments  = mail.Attachments

    unless attach_count == 0 then
      recur_save(attachments, save_dir, "save_attach: ")
    end
  rescue => ex
    puts "Error: save_attachment"
    puts ex
    exit 1
  end

private
  def folder(entry_id)
    @namespace.GetFolderFromID(entry_id)
  end

  def mail(entry_id)
    @namespace.GetItemFromID(entry_id)
  end

  def ole_call(ole_obj, method)
    if ole_obj.ole_respond_to?(method) then
      if method == "SentOn" then
        ole_obj.send(method)
      else
        ole_obj.send(method).encode("utf-8")
      end
    else
      puts "#{method} is undefined"
      if method == "SentOn" then
        Time.new
      else
        "unknown"
      end
    end
  end

  def recur_save(attachments, save_dir, message="save")
    attachments.each do |item|
      filename_utf8 = item.FileName.encode("utf-8")
      save_item_name = self.replace(filename_utf8)
      save_item = self.pathname(save_dir, save_item_name)

      item.SaveAsFile(save_item)
      puts "#{message}#{save_item_name}"

      # pick up zip in the .msg file
      if save_item_name =~ /^.*\.msg/ then
        attachments = @namespace.OpenSharedItem(save_item).Attachments
        recur_save(attachments, save_dir, " .msg unzip: -> ")
      end
    end
  end

  def recur_folders(folders)
    if folders.nil? then
      folders = @folders
      @f_index = 0
    else
      folders = folder(folders["entry_id"]).Folders
    end

    folders.each do |folder|
      @f_index += 1
      folder = { "name"     => folder.Name,
                 "count"    => folder.Items.Count,
                 "entry_id" => folder.EntryId }
      @folder_list[@f_index.to_s] = folder

      GC.start
      recur_folders(folder)
    end
    print "\r#{self.rjust(@f_index.to_s)} folders read..."
  rescue => ex
    puts "Error: folders"
    puts ex
    exit 1
  end
end

class String
  def show_encoding
    puts "-----------------"
    print "string         :"
    puts  self
    print "valid_encoding?:"
    puts  self.valid_encoding?
    print "encoding       :"
    puts self.encoding
    print "byte code      :"
    self.bytes {|b| print b.to_s + " "}
    puts ""
    puts "-----------------"
  end
end

if $0 == __FILE__ then
end
