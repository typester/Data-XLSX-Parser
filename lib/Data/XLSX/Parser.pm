package Data::XLSX::Parser;
use strict;
use warnings;

our $VERSION = '0.01';

use Data::XLSX::Parser::DocumentArchive;
use Data::XLSX::Parser::Workbook;
use Data::XLSX::Parser::SharedStrings;
use Data::XLSX::Parser::Styles;
use Data::XLSX::Parser::Sheet;

sub new {
    my ($class) = @_;

    bless {
        _row_event_handler => [],
        _archive           => undef,
        _workbook          => undef,
        _shared_strings    => undef,
    }, $class;
}

sub add_row_event_handler {
    my ($self, $handler) = @_;
    push @{ $self->{_row_event_handler} }, $handler;    
}

sub open {
    my ($self, $file) = @_;
    $self->{_archive} = Data::XLSX::Parser::DocumentArchive->new($file);
}

sub workbook {
    my ($self) = @_;
    $self->{_workbook} ||= Data::XLSX::Parser::Workbook->new($self->{_archive});
}

sub shared_strings {
    my ($self) = @_;
    $self->{_shared_strings} ||= Data::XLSX::Parser::SharedStrings->new($self->{_archive});
}

sub styles {
    my ($self) = @_;
    $self->{_styles} ||= Data::XLSX::Parser::Styles->new($self->{_archive});
}

sub sheet {
    my ($self, $sheet_id) = @_;
    $self->{_sheet} ||= Data::XLSX::Parser::Sheet->new($self, $self->{_archive}, $sheet_id);
}

sub _row_event {
    my ($self, $row) = @_;

    my $row_vals = [map { $_->{v} } @$row];
    for my $handler (@{ $self->{_row_event_handler} }) {
        $handler->($row_vals);
    }
}

1;

__END__

=head1 NAME

Data::XLSX::Parser - faster XLSX parser

=head1 SYNOPSIS

    use Data::Dumper;
    use Data::XLSX::Parser;
    
    my $parser = Data::XLSX::Parser->new;
    $parser->add_row_event_handler(sub {
        my ($row) = @_;
        print Dumper $row;
    });
    $parser->parse('foo.xlsx');

=head1 DESCRIPTION

Data::XLSX::Parser provides faster way to parse Microsoft Excel's .xlsx files.
The implementation of this module is highly inspired from Python's FastXLSX library.

This is SAX based parser, so you can parse very large XLSX file with lower memory usage.

=head1 THIS MODULE IS *ALPHA* QUALITY

This module is created for my current daily work that needs convert very huge excel file to csv, and perfectly work against my files but might not to all excel datas.

If you have some XSLX files that doesn't parse this module, please bug me with the files.

=head1 METHODS

=head2 new

Create new parser object.

=head2 add

=head1 AUTHOR

Daisuke Murase <typester@cpan.org>

=cut
