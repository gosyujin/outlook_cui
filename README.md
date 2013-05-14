# OutlookCui

## Installation

Add this line to your application's Gemfile:

    gem 'outlook_cui'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install outlook_cui

## Runtime Dependencies

- Application
  - Microsoft Office Outlook 2007
- Java (because parse *.msg file)
  - Java Runtime Environment 1.7.0
- Java home PATH
  - %JAVA_HOME%
- Jar files
  - msgparser (http://auxilii.com/msgparser/Page.php?id=16000)
    - poi (http://poi.apache.org/poifs/)
    - tnef (http://www.freeutils.net/source/jtnef/)
- Ruby Java Bridge
  - rjb(from bundle install)

if undefined JAVA_HOME

```
    Installing rjb (1.4.6)
    Gem::Installer::ExtensionBuildError: ERROR: Failed to build gem native extension
    .
    
            C:/Ruby193/bin/ruby.exe extconf.rb
    *** extconf.rb failed ***
    Could not create Makefile due to some reason, probably lack of
    necessary libraries and/or headers.  Check the mkmf.log file for more
    details.  You may need configuration options.
    
    Provided configuration options:
            --with-opt-dir
            --without-opt-dir
            --with-opt-include
            --without-opt-include=${opt-dir}/include
            --with-opt-lib
            --without-opt-lib=${opt-dir}/lib
            --with-make-prog
            --without-make-prog
            --srcdir=.
            --curdir
            --ruby=C:/Ruby193/bin/ruby
    extconf.rb:53:in `<main>': JAVA_HOME is not set. (RuntimeError)
```

## Usage

1. rake
1. Select a mail folder
1. Select download mails
1. Downloaded !!

and add task scheduler

### Rake args

- entry_id: folder's entry_id 
  - example: "00000000000000 ... 00000AEAAAD000"
      # download mail's entry_id
- id: mails's id
  - example: "1" or "1 3 9" or "1..10" or "all"
- save: saved path
  - default: ./mail
- verbose: show detail
  - default: false

```
    $ rake entry_id="00000000000000 ... 00000AEAAAD000" id="all" save="C:/mail"
```

### Example

- rake

```
    $ rake
      1|         0|メールボックス
      2|         0|削除済みアイテム
      3|         1|受信トレイ
      4|         0|hoge
      5|       818|mail
      6|         2|spam
    -------------------------------
     id|mail count|folder name
```

- Select a mail folder and show folder's mails

```
    Select a folder's id:
    5
      1|                    0|test について
      2|                    2|RE: test について
      3|                    1|(添付)Excel方眼紙
    -------------------------------------------
     id|attachment file count|mail subject
```

- Select download mails

```
    Select mails's ids(1 or 1 2 3 or 1..3 or all):
    1..3
```

- Downloaded !!

```
    save_mail  : testについて       # download mail id = 1
    save_mail  : RE: testについて   # download mail id = 2
    save_attach: test.txt           #  -> mail's attachment file
    save_attach: hogehoge.msg       #  -> mail's attachment file(.msg)
    .msg unzip : -> 中身.zip        #  -> parse .msg file
    save_mail  : (添付)Excel方眼紙  # download mail id = 3
    save_attach: Excel方眼紙.zip    #  -> mail's attachment file
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
