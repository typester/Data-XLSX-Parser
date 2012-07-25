use strict;
use warnings;
use utf8;
use FindBin;

use Test::More;

use_ok 'Data::XLSX::Parser';

my $parser = Data::XLSX::Parser->new;
isa_ok $parser, 'Data::XLSX::Parser';

$parser->open("$FindBin::Bin/../private-data-20120717.xlsx");

my $shared_strings = $parser->shared_strings;

is $shared_strings->count, 8161, 'count ok';

is $shared_strings->get(1), '問題文', 'get 1 ok';
is $shared_strings->get(1000), 'リオデジャネイロ', 'get 1000 ok';


done_testing;
