package Data::XLSX::Parser;
use strict;
use warnings;

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
