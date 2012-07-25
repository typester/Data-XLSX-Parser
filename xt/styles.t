use strict;
use warnings;
use utf8;
use FindBin;

use Test::More;

use_ok 'Data::XLSX::Parser';

my $parser = Data::XLSX::Parser->new;
isa_ok $parser, 'Data::XLSX::Parser';

$parser->open("$FindBin::Bin/../private-data-20120717.xlsx");

my $styles = $parser->styles;



done_testing;
