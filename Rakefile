# -*- encoding: utf-8 -*-
require "bundler/gem_tasks"

$LOAD_PATH.unshift(File.expand_path('../lib', __FILE__))
require 'outlook_cui'
require 'pp'

include Outlook::Utility

task :default => :outlook

entry_id = ""
ids = []
mails = nil

desc 'DEFAULT args entry_id="000000000" id="all" save="C:/hogehoge"'
task :outlook do
  # folder's entry_id
  # ENV["entry_id"] = "00000000000000 ... 00000AEAAAD000"
  # download mail's entry_id
  # ENV["id"] = "all"
  # download path # fix me Japanese pathname
  # ENV["save"]     = "./mail"
  Rake::Task['save'].invoke
end

task :folders do
  if ENV["entry_id"].nil?
    folders = OutlookCui.folders
    # pp folders

    folders.each do |id, folder|
      puts "#{self.rjust(id)}|" \
           "#{self.rjust(folder["count"])}|" \
           "#{folder["name"]}"
      # puts "#{folder["entry_id"]}"
    end
    puts "Choose a folder's id:"
    id = STDIN.gets.chomp!

    entry_id = folders[id]["entry_id"]
  else
    entry_id = ENV["entry_id"]
  end
end

task :mails => [:folders] do
  mails = OutlookCui.mails(entry_id)
  # pp mails

  if ENV["id"].nil?
    mails.each do |id, mail|
      puts "#{self.rjust(id)}|" \
           "#{self.rjust(mail["attach_count"])}|" \
           "#{mail["subject"]}"
      # puts "#{mail["entry_id"]}"
    end
    puts "Choose mails's ids(1 or 1 2 3 or 1..3 or all):"
    id = STDIN.gets.chomp!
  else
    id = ENV["id"]
  end

  case id
  when "all"
    ids = Range.new("1", mails.length.to_s)
  when /^[0-9]*\.\.[0-9]*$/
    # Range ex: 10..23
    range = id.split("..")
    ids = Range.new(range[0], range[-1])
  else
    # Array and other ex: 1 or 1 2 3
    ids = id.split(" ")
  end
end

task :save => [:mails] do
  ids.each do |id|
    OutlookCui.save_mail(mails[id]["entry_id"], ENV["save"], true)
  end 
end
