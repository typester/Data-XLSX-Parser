package Data::XLSX::Parser::SharedStrings;
use strict;
use warnings;

use XML::Parser::Expat;
use Archive::Zip ();
use File::Temp;

sub new {
    my ($class, $archive) = @_;

    my $self = bless {
        _data      => [],

        _is_string => 0,
        _buf       => '',
    }, $class;

    my $fh = File::Temp->new( SUFFIX => '.xml' );

    my $handle = $archive->shared_strings or return $self;
    die 'Failed to write temporally file: ', $fh->filename
        unless $handle->extractToFileNamed($fh->filename) == Archive::Zip::AZ_OK;

    my $parser = XML::Parser::Expat->new;
    $parser->setHandlers(
        Start => sub { $self->_start(@_) },
        End   => sub { $self->_end(@_) },
        Char  => sub { $self->_char(@_) },
    );
    $parser->parse($fh);

    $self;
}

sub count {
    my ($self) = @_;
    scalar @{ $self->{_data} };
}

sub get {
    my ($self, $index) = @_;
    $self->{_data}->[$index];
}

sub _start {
    my ($self, $parser, $name, %attrs) = @_;
    $self->{_is_string} = 1 if $name eq 'si';
}

sub _end {
    my ($self, $parser, $name) = @_;
    $self->{_is_string} = 0;

    if ($name eq 'si') {
        push @{ $self->{_data} }, $self->{_buf};
        $self->{_buf} = '';
    }
}

sub _char {
    my ($self, $parser, $data) = @_;
    $self->{_buf} .= $data if $self->{_is_string};
}

1;
