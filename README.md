# AssOle

Provides features for easy access to 1C:Enterprise Ole from Ruby code.
Main things of this gem is `AssOle::Runtimes` and `AssOle::Snippets`

`AssOle::Runtimes` provides features for control, despatching and easy access to
the 1C:Enterprise Ole connectors. `AssOle::Runtimes` inclides mixins which
provides `ole_connector` method returned specified Ole connector.

`AssOle::Snippets` provides features for transparent access to 1C:Enterprise Ole
methods and properties from Ruby objects like as they are was own Ruby object
methods. In other words `AssOle::Snippetes` forvarding call unknown methods
to the `ole_connector` in `method_missing` handler.

Both this things makes Ruby code shorter and tidier

## Examples

TODO

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'ass_ole'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install ass_ole

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/ass_ole.

