package Data::XLSX::Parser::Sheet;
use strict;
use warnings;

use File::Temp;
use XML::Parser::Expat;
use Archive::Zip ();

use constant {
    STYLE_IDX          => 'i',
    STYLE              => 's',
    FMT                => 'f',
    REF                => 'r',
    COLUMN             => 'c',
    VALUE              => 'v',
    TYPE               => 't',
    TYPE_SHARED_STRING => 's',
    GENERATED_CELL     => 'g',
};

sub new {
    my ($class, $doc, $archive, $sheet_id) = @_;

    my $self = bless {
        _document => $doc,

        _data => '',
        _is_sheetdata => 0,
        _row_count => 0,
        _current_row => [],
        _cell => undef,
        _is_value => 0,

        _shared_strings => $doc->shared_strings,
        _styles         => $doc->styles,

    }, $class;

    my $fh = File::Temp->new( SUFFIX => '.xml' );

    my $handle = $archive->sheet($sheet_id);
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

sub _start {
    my ($self, $parser, $name, %attrs) = @_;

    if ($name eq 'sheetData') {
        $self->{_is_sheetdata} = 1;
    }
    elsif ($self->{_is_sheetdata} and $name eq 'row') {
        $self->{_current_row} = [];
    }
    elsif ($name eq 'c') {
        $self->{_cell} = {
            STYLE_IDX() => $attrs{ STYLE_IDX() },
            TYPE()      => $attrs{ TYPE() },
            REF()       => $attrs{ REF() },
            COLUMN()    => scalar(@{ $self->{_current_row} }) + 1,
        };
    }
    elsif ($name eq 'v') {
        $self->{_is_value} = 1;
    }
}

sub _end {
    my ($self, $parser, $name) = @_;

    if ($name eq 'sheetData') {
        $self->{_is_sheetdata} = 0;
    }
    elsif ($self->{_is_sheetdata} and $name eq 'row') {
        $self->{_row_count}++;
        $self->{_document}->_row_event( delete $self->{_current_row} );
    }
    elsif ($name eq 'c') {
        my $c = $self->{_cell};
        $self->_parse_rel($c);

        if (($c->{ TYPE() } || '') eq TYPE_SHARED_STRING()) {
            my $idx = int($self->{_data});
            $c->{ VALUE() } = $self->{_shared_strings}->get($idx);
        }
        else {
            $c->{ VALUE() } = $self->{_data};
        }

        $c->{ STYLE() } = $self->{_styles}->cell_style( $c->{ STYLE_IDX() } );
        $c->{ FMT() }   = my $cell_type =
            $self->{_styles}->cell_type_from_style($c->{ STYLE() });

        my $v = $c->{ VALUE() };
        if (defined $v and $c->{ FMT() } =~ /^datetime\.(date)?(time)?$/) {
            # datetime
            warn 'datetime';
        }
        else {
            if (!defined $v) {
                $c->{ VALUE() } = '';
            }
            elsif ($cell_type ne 'unicode') {
                warn 'not unicode';
                $c->{ VALUE() } = $v;
            }
        }

        push @{ $self->{_current_row} }, $c;

        $self->{_data} = '';
        $self->{_cell} = undef;
    }
    elsif ($name eq 'v') {
        $self->{_is_value} = 0;
    }
}

sub _char {
    my ($self, $parser, $data) = @_;

    if ($self->{_is_value}) {
        $self->{_data} .= $data;
    }
}

sub _parse_rel {
    my ($self, $cell) = @_;

    my ($column, $row) = $cell->{ REF() } =~ /([A-Z]+)(\d+)/;

    my $v = 0;
    my $i = 0;
    for my $ch (split '', $column) {
        my $s = length($column) - $i++ - 1;
        $v += (ord($ch) - ord('A') + 1) * (26**$s);
    }

    $cell->{ REF() } = [$v, $row];

    if ($cell->{ COLUMN() } > $v) {
        die sprintf 'Detected smaller index than current cell, something is wrong! (row %s): %s <> %s', $row, $v, $cell->{ COLUMN() };
    }
}

1;


