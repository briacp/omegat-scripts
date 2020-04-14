#!perl
use strict;
use XML::Writer;
use Spreadsheet::XLSX;
use IO::File;
use Getopt::Long;
use Pod::Usage;
use Pod::Text;
use Tk;
use Encode qw(decode);

use constant VERSION => "1.0.3";

# pp --gui -o c:\users\briac\desktop\xlsx2tmx.exe -M Pod::Text xlsx2tmx.pl

my (
    $help,        $man,             $verbose,
    $source_lang, $target_lang,     $source_col,
    $target_col,  $worksheet_index, $encoding
) = ( undef, undef, 0, "EN", "FR", 1, 2, 1, "utf-8" );

my $segtype = "paragraph";    # "sentence"

my $xlsx_file;

if (@ARGV) {

    GetOptions(
        "source=s"     => \$source_lang,
        "target=s"     => \$target_lang,
        "worksheet=i"  => \$worksheet_index,
        "col-target=i" => \$target_col,
        "col-source=i" => \$source_col,
        "encoding=s"   => \$encoding,
        "verbose"      => \$verbose,
        "help|?"       => \$help,
        "man"          => \$man
    ) or pod2usage(2);

    pod2usage(1) unless @ARGV;

    pod2usage(1) if $help;
    pod2usage( -exitstatus => 0, -verbose => 2 ) if $man;

    $xlsx_file = $ARGV[0];

    generate_tmx();
}
else {
    my ( $MW, %ZWIDGETS, );

    $verbose = 1;

    $MW = MainWindow->new( -title => "Excel to TMX", -padx => 10, -pady => 10 );

    ZloadImages();
    ZloadFonts();

    my $row = 0;

    # Widget label_title isa Label
    $ZWIDGETS{'label_title'} = $MW->Label(
        -font      => 'MS_Sans_Serif_12_bold_roman_',
        -takefocus => 0,
        -text      => 'Excel to TMX',
      )->grid(
        -row        => $row,
        -column     => 0,
        -columnspan => 3,
      );

    # Widget label_source isa Label
    $ZWIDGETS{'label_source'} = $MW->Label(
        -takefocus => 0,
        -text      => 'Source language :',
      )->grid(
        -row    => ++$row,
        -column => 0,
        -sticky => 'e',
      );

    # Widget entry_source isa Entry
    $ZWIDGETS{'entry_source'} = $MW->Entry(
        -exportselection => 1,
        -textvariable    => \$source_lang,
        -width           => 7,
      )->grid(
        -row    => $row,
        -column => 1,
        -sticky => 'w',
      );

    # Widget label_target isa Label
    $ZWIDGETS{'label_target'} = $MW->Label(
        -takefocus => 0,
        -text      => 'Target Language :',
      )->grid(
        -row    => ++$row,
        -column => 0,
        -sticky => 'e',
      );

    # Widget entry_target isa Entry
    $ZWIDGETS{'entry_target'} = $MW->Entry(
        -exportselection => 1,
        -textvariable    => \$target_lang,
        -width           => 7,
      )->grid(
        -row    => $row,
        -column => 1,
        -sticky => 'w',
      );

    # Widget label_xls_file isa Label
    $ZWIDGETS{'label_xls_file'} = $MW->Label(
        -takefocus => 0,
        -text      => 'Excel file :',
      )->grid(
        -row    => ++$row,
        -column => 0,
        -sticky => 'e',
      );

    # Widget entry_xlsx_file isa Entry
    $ZWIDGETS{'entry_xlsx_file'} = $MW->Entry(
        -exportselection => 1,
        -width           => 50,
        -textvariable    => \$xlsx_file,
      )->grid(
        -row    => $row,
        -column => 1,
        -sticky => 'w',
      );

    # Widget Button1 isa Button
    $ZWIDGETS{'button_filedir'} = $MW->Button(
        -text    => 'Choose...',
        -command => sub {
            $xlsx_file = $MW->getOpenFile(
                -filetypes => [ [ "Excel Files", ".xlsx" ], [] ], );
        },

      )->grid(
        -row    => $row,
        -column => 2,
      );

    $ZWIDGETS{'label_empty'} = $MW->Label( -takefocus => 0, )
      ->grid( -row => ++$row, -column => 0, -columnspan => 4 );

    #==========================================================
    $ZWIDGETS{'label_worksheet_index'} = $MW->Label(
        -takefocus => 0,
        -text      => 'Worksheet # :',
      )->grid(
        -row    => ++$row,
        -column => 0,
        -sticky => 'e',
      );

    $ZWIDGETS{'entry_worksheet_index'} = $MW->Entry(
        -exportselection => 1,
        -textvariable    => \$worksheet_index,
        -width           => 3,
      )->grid(
        -row    => $row,
        -column => 1,
        -sticky => 'w',
      );

    $ZWIDGETS{'label_source_col'} = $MW->Label(
        -takefocus => 0,
        -text      => 'Source Column # :',
      )->grid(
        -row    => ++$row,
        -column => 0,
        -sticky => 'e',
      );

    $ZWIDGETS{'entry_source_col'} = $MW->Entry(
        -exportselection => 1,
        -textvariable    => \$source_col,
        -width           => 3,
      )->grid(
        -row    => $row,
        -column => 1,
        -sticky => 'w',
      );

    $ZWIDGETS{'label_target_col'} = $MW->Label(
        -takefocus => 0,
        -text      => 'Target Column # :',
      )->grid(
        -row    => ++$row,
        -column => 0,
        -sticky => 'e',
      );

    $ZWIDGETS{'entry_target_col'} = $MW->Entry(
        -exportselection => 1,
        -textvariable    => \$target_col,
        -width           => 3,
      )->grid(
        -row    => $row,
        -column => 1,
        -sticky => 'w',
      );

    #==========================================================
    $ZWIDGETS{'label_empty2'} = $MW->Label( -takefocus => 0, )
      ->grid( -row => ++$row, -column => 0, -columnspan => 4 );

    # Widget button_generate isa Button
    $ZWIDGETS{'button_generate'} = $MW->Button(
        -command => \&generate_tmx,
        -text    => '  Generate TMX  ',
        -padx    => 4,
        -pady    => 4,
      )->grid(
        -row        => ++$row,
        -column     => 0,
        -columnspan => 4,
      );
    $ZWIDGETS{'label_empty3'} = $MW->Label( -takefocus => 0, )
      ->grid( -row => ++$row, -column => 0, -columnspan => 4 );

    $ZWIDGETS{'log'} = $MW->Scrolled(
        'Text',
        -scrollbars => 'e',
        -padx       => 4,
        -pady       => 4,
        -height     => 10,
        -wrap       => 'none',
      )->grid(
        -row        => ++$row,
        -column     => 0,
        -columnspan => 4,
      );
    my $text = $ZWIDGETS{'log'}->{SubWidget}->{text};
    tie *STDOUT, ref $text, $text;

    $ZWIDGETS{'label_info'} = $MW->Label(
        -anchor    => 'e',
        -font      => 'MS_Sans_Serif_6_normal_roman_',
        -padx      => 4,
        -pady      => 4,
        -takefocus => 0,
        -text      => "xlsx2tmx v." . VERSION,
      )->grid(
        -row        => ++$row,
        -column     => 0,
        -columnspan => 3,
        -sticky     => 'e',
      );

    MainLoop();

    sub ZloadImages {
    }

    sub ZloadFonts {
        $MW->fontCreate(
            'MS_Sans_Serif_6_normal_roman_',
            -weight     => 'normal',
            -underline  => 0,
            -family     => 'MS Sans Serif',
            -slant      => 'roman',
            -size       => -11,
            -overstrike => 0,
        );
        $MW->fontCreate(
            'MS_Sans_Serif_12_bold_roman_',
            -weight     => 'bold',
            -underline  => 0,
            -family     => 'MS Sans Serif',
            -slant      => 'roman',
            -size       => -18,
            -overstrike => 0,
        );
    }
}

sub generate_tmx {

    ( my $xml_file = $xlsx_file ) =~
      s/\.xlsx$/-${source_lang}_$target_lang.tmx/;

    print "Excel file: $xlsx_file\n" if $verbose;
    print "TMX   file: $xml_file\n"  if $verbose;

    if ( !-e $xlsx_file ) {
        print "XLSX file $xlsx_file does not exist.\n";
        return;
    }

    my $workbook = Spreadsheet::XLSX->new($xlsx_file);

    if ( !defined $workbook ) {
        print "Cannot open XLSX file $xlsx_file.\n";
        return;
    }

    my $output = new IO::File("> $xml_file");
    binmode $output, ":utf8";
    my $writer = new XML::Writer( OUTPUT => $output );
    $writer->xmlDecl("UTF-8");
    $writer->doctype( "tmx", undef, "tmx11.dtd" );
    $writer->startTag( "tmx", version => "1.1" );

    $writer->startTag(
        "header",
        creationtool        => "xlsx2tmx",
        creationtoolversion => VERSION,
        segtype             => $segtype,
        "o-tmf"             => "OmegaT TMX",
        "adminlang"         => "EN-US",
        srclang             => $source_lang,
        datatype            => "plaintext"
    );
    $writer->endTag("header");
    $writer->startTag("body");

    my $worksheet = $workbook->worksheet( $worksheet_index - 1 );
    my ( $row_min, $row_max ) = $worksheet->row_range();

    print "Found $row_min/$row_max rows in worksheet $worksheet_index\n";

    for my $row ( $row_min .. $row_max ) {
        print "Row $row\n" if $verbose;

        my $source_cell = $worksheet->get_cell( $row, $source_col - 1 );
        next unless $source_cell;

        my $source = $source_cell->value;
        my $target_cell = $worksheet->get_cell( $row, $target_col - 1 );

        next unless $target_cell;
        my $target = $target_cell->value;

        if (
                     $source ne $target
                     &&
            $target ne ""
          )
        {
            print "\t$source\n\t$target\n" if $verbose;
            $writer->startTag("tu");
            $writer->startTag( "tuv", lang => $source_lang );
            $writer->startTag("seg");
            $writer->characters( trim($source) );
            $writer->endTag("seg");
            $writer->endTag("tuv");
            $writer->startTag( "tuv", lang => $target_lang );
            $writer->startTag("seg");
            $writer->characters( trim($target) );
            $writer->endTag("seg");
            $writer->endTag("tuv");
            $writer->endTag("tu");
        }

    }

    $writer->endTag("body");
    $writer->endTag("tmx");
    $writer->end();

    print "TMX created : $xml_file\n";
}

sub trim {
    my $s = shift;
    $s = decode( $encoding, $s );
    $s =~ s/^\s+|\s+$//g;

    $s;
}

__END__

=head1 NAME

xlsx2tmx - TMX generation from Excel .xlsx file

=head1 SYNOPSIS

xlsx2tmx [options] file.xlsx

 Options:
    --source         Source language (default "EN-US")
    --target         Target language (default "FR-FR")
    --worksheet      Index of the Excel worksheet (default 1)
    --col-source     Index of the source column (default 1)
    --col-target     Index of the target column (default 2)
    --encoding       Encoding to use (default "utf-8")
    --verbose        Verbose mode
    --help
    --man

    briacp@gmail.com

=cut
