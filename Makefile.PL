use strict;
use warnings;

BEGIN {
    my @devmods = qw(
        Module::Install::AuthorTests
        Module::Install::ReadmeFromPod
        Module::Install::Repository
    );
    my @not_available;

    eval qq{use inc::Module::Install; 1;} or push @not_available, 'inc::Module::Install';
    for my $mod (@devmods) {
        eval qq{require $mod} or push @not_available, $mod;
    }
    if (@not_available) {
        print qq{# The following modules are not available.\n};
        print qq{# `$^X $0 | cpanm` will install them:\n};
        print $_, "\n" for @not_available;
        print "\n";
        exit -1;
    }
}

use inc::Module::Install;

name 'Data-XLSX-Parser';
all_from 'lib/Data/XLSX/Parser.pm';

readme_from 'lib/Data/XLSX/Parser.pm';
author_tests 'xt';
auto_set_repository;

requires 'Archive::Zip';
requires 'XML::Parser::Expat';
requires 'File::Temp';

WriteAll;
