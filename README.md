# XlsxWriter

This little library is a fast XLSX files writer, base on quite awesome [libxlsxwriter](https://libxlsxwriter.github.io/) written in C. Unlike existing solutions, such as [RubyXL](https://github.com/weshatheleopard/rubyXL) is does not let you manipulate existing files or read them. A sheer goal of XlsxWriter is to be able to dump data into an Excel file and send it to whoever requested it.

This is very work-in-progress now and supports only basic things I needed for tests. Feel free to add more or open an issue if you need something more from libxlsxwriter.

## Installation

As a preparation step, you need to install libxlsxwriter.

* On Mac apparently it's enough to `brew install libxlsxwriter` (not tested)
* On Linux it's usual `git clone https://github.com/jmcnamara/libxlsxwriter.git` + `make` + `make install` (unless you're on Arch, then you have packages in AUR available)
* On Windows you're on your own, but please let me know if you succeed, so I can update this README

Add this line to your application's Gemfile:

```ruby
gem 'xlsxwriter', github: 'katafrakt/ruby-xlsxwriter'
```

And then execute:

    $ bundle

## Usage

Basic usage:

```ruby
XlsxWriter.create('test.xlsx') do |excel|
  excel.set_column_styles([
                            { width: 10 },
                            { width: 20 },
                            { width: 12, copy_for_next: 7 },
                            { width: nil, copy_for_next: 1 }, # default width
                            { width: 10 }
                          ])
  excel.write_row(['The Title'], format: { bold: true, font_size: 25, height: 27 })
  excel.skip_row
  excel.write_row [nil, 2, 'test', Class.new, 2.67]
  excel.write_row [nil, 2, 'test', Class.new, 2.67], offset: 5
end
```

You also have memory-efficient mode available. To enable it, initialize like this:

```
XlsxWriter.create('test.xlsx', constant_memory: true)
```

In my local tests with rewriting 550k-rows CSV file into XLSX it went down from 1140MB peak memory usage to just 153MB.

Caveats:
* You need to write your rows in orded (it's impossible to do otherwise using public API, but you could do it using low-level `XlsxWriter::C` bindings).
* It uses something called string-inlining, which work fine in most spreadsheet software, with expected exception of Apple Numbers, which shows empty cells instead. You're gonna need to tell people to use real software instead (which is a good piece of advise anyway).

More details: http://libxlsxwriter.github.io/working_with_memory.html

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/katafrakt/ruby-xlsxwriter.

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).
