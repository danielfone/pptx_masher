# PPTXMasher

This (unpublished) gem is for merging slides from Powerpoint presentations.
It also allows you to make small modifications to the slides.

It has been developed for a specific in-house application,
so it is only really supported in that context.
This is not under active development except for the needs of the original application,
but the code is published here in case it's any use to anyone.

Some parts were heavily influenced by the [powerpoint gem](https://github.com/pythonicrubyist/powerpoint).

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'pptx_masher'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install pptx_masher

## Usage

    p = PPTXMasher::Presentation.load "path/to/presentation.pptx"
    slide = p.slides[1]
    slide.replace_text "[title]", "acutal title"
    slide.replace_media 'image3.jpeg', 'path/to/other/image/jpeg'
    p.save "path/to/updated/presentation.pptx"

    p.add_slide slide

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release` to create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

1. Fork it ( https://github.com/[my-github-username]/pptx_masher/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
