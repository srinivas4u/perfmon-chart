#!/usr/bin/perl 
# Script to convert Perform data in csv for to Excel Charts

use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel;

# intial configuration

my $dataloc = 70;
chomp($dataloc);
my $plotchart = $dataloc + 2;

print "Enter the Perfmom file in csv format : ";
my $inputfile = <> ;

chomp ($inputfile);
my $workbook = new Spreadsheet::WriteExcel("$inputfile.xls");

my $chart1 = $workbook->add_chart( type => 'line', embedded => 1 );

# Reading file for input of Excel sheet
open(FR, "<$inputfile.csv") || die "Can't open $inputfile.csv: $!";

@line =<FR>;

	$y = scalar(@line);
		@hostName = split(/\",\"/, $line[0]);
			@StartTime = split(/\",\"/, $line[1]);
				@EndTime = split(/\",\"/, $line[$y-1]);
					$StartTime[0] =~ s/\"//igo;
			$EndTime[0] =~ s/\"//igo;
		$hostName[1] =~ s/\\//igo;
	$hostName[1] =~ s/M(.+)//igo;

close(FR);
	
	print "$hostName[1]\n"; 
		print "Start Time $StartTime[0] \n";
	print "End Time $EndTime[0] \n";

my $sheet = $workbook->add_worksheet("$hostName[1]");

my $boldFormat = $workbook->add_format();

$boldFormat->set_bold();
my $date_format = 'hh:mm:ss';

my $Timeformat =  $workbook->add_format(
                                        num_format => $date_format,
                                        align      => 'center'
                                       );
my $date;

my $row=$dataloc;

my @longest=(0,0,0,0,0,0,0);

open(FW, "<$inputfile.csv") || die "Can't open $inputfile.csv: $!";

while(<FW>) {

     chomp;

        my @fields = split(/,/, $_);


        for(my $col = 0; $col < @fields; $col++) {


						$fields[$col] =~ s/\"//;
						$fields[$col] =~ s/\"//;					
					#	$fields[$col] =~ s/(.+?)\/(.+?)\/(.+?)\//;
					#	$fields[$col] =~ s/\.(.+?)//;


             if ($row == $dataloc ) {


                     $sheet->write($row, $col, $fields[$col], $boldFormat);

             } else {

                     if ($col == 0) {
						 		
							 $fields[$col] = substr($fields[$col],10,9);	
                             
							 $sheet->write($row, $col,

							 $fields[$col],$Timeformat);
												 
                     } else {
						
                             $sheet->write($row, $col, $fields[$col]);

                     }

             }

             if ($longest[$col] < length($fields[$col])) {

                     $longest[$col] = length($fields[$col]);
             }

        }

     $row++;

}

$longest[2]+=3;

$longest[3]+=3;

for(my $i = 0; $i < @longest; $i++) {

     $sheet->set_column($i,$i,$longest[$i]);

}

# Configure the series.

close(FW);

#Start Time and End Time
$sheet->write  (1, 0, 'Start Time', $boldFormat);
$sheet->write  (1, 1, ''.$StartTime[0].'', $boldFormat);
$sheet->write  (2, 0, 'End Time', $boldFormat);
$sheet->write  (2, 1, ''.$EndTime[0].'', $boldFormat);

#Chart 1 Memory						

$chart1->add_series(
    categories => '='.$hostName[1].'!$A$72:$A$'.$row.'',
    values     => '='.$hostName[1].'!$B$72:$B$'.$row.'',
    name       => '% Committed Bytes In Use',
);



# Add some labels.
$chart1->set_title( name => 'Results of Memory analysis' );
$chart1->set_x_axis( name => 'Time in hh:mm:ss' );
$chart1->set_y_axis( name => '% Committed Bytes In Use' );

# Insert the chart into the main worksheet.
$sheet->insert_chart( 'A4', $chart1 );

#Chart 2 IO Read

my $chart2 = $workbook->add_chart( type => 'line', embedded => 1 );

$chart2->add_series(
    categories => '='.$hostName[1].'!$A$72:$A$'.$row.'',
    values     => '='.$hostName[1].'!$C$72:$C$'.$row.'',
    name       => 'IO Read Operations/sec',
);

# Add some labels.
$chart2->set_title( name => 'Results of IO Read analysis' );
$chart2->set_x_axis( name => 'Time in hh:mm:ss' );
$chart2->set_y_axis( name => 'IO Read Operations/sec' );

# Insert the chart into the main worksheet.
$sheet->insert_chart( 'C25', $chart2 );

#Chart 3 CPU Usage 

my $chart3 = $workbook->add_chart( type => 'line', embedded => 1 );

$chart3->add_series(
    categories => '='.$hostName[1].'!$A$72:$A$'.$row.'',
    values     => '='.$hostName[1].'!$E$72:$E$'.$row.'',
    name       => '% Processor Time',
);

# Add some labels.
$chart3->set_title( name => 'Results of CPU analysis' );
$chart3->set_x_axis( name => 'Time in hh:mm:ss' );
$chart3->set_y_axis( name => '% Processor Time' );

# Insert the chart into the main worksheet.
$sheet->insert_chart( 'C4', $chart3 );

#Chart 4 IO Write 

my $chart4 = $workbook->add_chart( type => 'line', embedded => 1 );

$chart4->add_series(
    categories => '='.$hostName[1].'!$A$72:$A$'.$row.'',
    values     => '='.$hostName[1].'!$D$72:$D$'.$row.'',
    name       => 'IO Write Operations/sec',
);

# Add some labels.
$chart4->set_title( name => 'Results of IO Write Operations analysis' );
$chart4->set_x_axis( name => 'Time in hh:mm:ss' );
$chart4->set_y_axis( name => 'IO Write Operations/sec' );

# Insert the chart into the main worksheet.
$sheet->insert_chart( 'A25', $chart4 );

$workbook->close();    
