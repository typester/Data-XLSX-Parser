use strict;
use warnings;
use utf8;

use FindBin;
use Test::More;

use Data::XLSX::Parser;

my $parser = Data::XLSX::Parser->new;

my $fn = __FILE__;
$fn =~ s{t$}{xlsx};

$parser->open($fn);

$parser->add_row_event_handler(sub {
    my ($row) = @_;

    use YAML;
    warn Dump $row;
});

$parser->sheet(1);

done_testing;
