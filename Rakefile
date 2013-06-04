# -*- encoding: utf-8 -*-
require "bundler/gem_tasks"

$LOAD_PATH.unshift(File.expand_path('../lib', __FILE__))
require 'outlook_cui'
require 'pp'

include Outlook::Utility

task :default => :outlook

# ENVs
# entry_id:  folder's entry_id.
#            ex
#               rake entry_id="00000000000000 ... 00000AEAAAD000"
#
# id:        download mail's id.
#            ex
#               rake id="all"       # select ALL mails
#               rake id="2..30"     # select id = 2, 3 .. 29, 30 mails
#               rake id="1 3 4 5 6" # select id = 1, 3, 4, 5, 6 mails
#
# save:      download save path. (fix me Japanese pathname)
#            ex
#               rake save="./mail"
#
# limit:     show and download mail, set limit.
#            ex
#               rake limit=20
#
# attach:    show mail exist attachment file. (except nil!)
#            ex
#               rake                # show ALL
#               rake attach="true"  # attachment only
#               rake attach="yes"   # attachment only
#               rake attach="on"    # attachment only
#               rake attach=        # attachment only!
#               rake attach="no"    # attachment only!
#               rake attach="false" # attachment only!
#
# verbose:   show entry_id etc. (except nil!)
#            ex
#               same as above "attach"
#               rake                # verbose OFF
#               rake verbose="true" # verbose on
#               rake verbose="no"   # verbose on!
#               etc.

verbose = !verbose      unless ENV['verbose'].nil?

entry_id = ""
ids = []
mails = nil

task :default => :outlook

desc 'DEFAULT args entry_id id save limit attach verbose'
task :outlook do
  Rake::Task['save'].invoke
end

task :folders do
  if ENV["entry_id"].nil?
    folders = OutlookCui.folders

    # require 'pp'
    # pp folders
    ## sort
    # folders = folders.sort_by{|id, folder| folder["name"]}
    # pp folders

    folders.each do |id, folder|
      out = ""
      out << "#{folder["entry_id"]}|" if verbose
      out << "#{self.rjust(id)}|" \
             "#{self.rjust(folder["count"])}|" \
             "#{folder["name"]}"
      puts out
    end
    puts "Choose a folder's id:"

    while(0)
      id = STDIN.gets.chomp!
      # "a".to_i #=> 0
      if id.to_i > 0 and id.to_i <= folders.length then
        break
      else
        puts "Out of range or if you input string."
        puts "One more."
      end
    end

    entry_id = folders[id]["entry_id"]
  else
    entry_id = ENV["entry_id"]
  end
end

task :mails => [:folders] do
  mails = OutlookCui.mails(entry_id)

  if ENV["id"].nil?
    mails.each do |id, mail|
      out = ""
      out << "#{mail["entry_id"]}|" if verbose
      out << "#{self.rjust(id)}|" \
             "#{self.rjust(mail["attach_count"])}|" \
             "#{mail["sent"]}|" \
             "#{mail["subject"]}"
      puts out
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
