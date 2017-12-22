[![Gem Version](https://badge.fury.io/rb/ass_ole.svg)](https://badge.fury.io/rb/ass_ole)
# AssOle

Provides features for easy access to 1C:Enterprise Ole from Ruby code.
Main things of this gem is `AssOle::Runtimes` and `AssOle::Snippets`

`AssOle::Runtimes` provides features for control, despatch and easy access to
the 1C:Enterprise Ole connectors. `AssOle::Runtimes` inclides mixins which
provides `ole_connector` method returned specified Ole connector.

`AssOle::Snippets` provides features for transparent access to 1C:Enterprise Ole
methods and properties from Ruby objects like as they are was own Ruby object
methods. In other words `AssOle::Snippetes` forvarding call unknown methods
to the `ole_connector` in the `method_missing` handler.

Both this things makes Ruby code shorter and tidier

## Attention

`AssOle::Runtimes` closes all ole connections in `at_exit` hook. You should
checks order of modules loading.
For example if `ass_ole` uses with `minitest` first load
`ass_ole` secont load `minitest` otherwise all ole connections will be closed
before start tests executing:

```ruby
requre 'ass_ole'
requre 'minitest/autorun'
```

## Examples

More about it and how to use see [test/examples_test.rb](test/examples_test.rb)

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

## Testing

    $ export SIMPLECOV=YES && rake test

## Contributing

Bug reports and pull requests are welcome on GitHub at
https://github.com/leoniv/ass_ole.

