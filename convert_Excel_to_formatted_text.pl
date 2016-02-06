#!/usr/bin/perl -w
# Author : Abraham Taylor
# Date Last Revised : 02.05.2016
# This script is for use by King County Archives
# INPUT : This script takes in an excel file. 
# INPUT NOTES: There needs to be at least four worksheets. Only the first four worksheets will be processed.
# INPUT NOTES: The first worksheet of the excel file should be the header. It must have at least three columns. Only the first three columns will be processed. The first column is descriptive information (which is ignored by this script). The second column should be the field names for the output (this is where REC_ID should go), and the third column should be the values.
# INPUT NOTES: The second, third, and fourth worksheets are processed in the same way. The first row of the these sheets should be descriptive information (which is ignored by this script). The second row should be the field names for the output. The following rows should be the actual data.
# OUTPUT : The first part of the output will be the header worksheet. Then the next three worksheets will be outputted. The format of the output is that each line will have the field name followed by the value. There will be a few spaces between each worksheet's contents in the output.

use File::Find;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw/ :LOCALE :DATE /;

$Win32::OLE::Warn = 3;  # die on errors...

#read in file names
print "Please enter the FULL path to the Excel file to be read as input: ";
$excelFile = <STDIN>;
chomp($excelFile);

print "Please enter the FULL path to the output file: ";
$outfile = <STDIN>;
chomp ($outfile);

#these three lines are for debugging purposes
#$excelFile = "C:\\Users\\abraham\\Desktop\\A14-074_Perl.xlsx";
#$outfile = "output.txt";

print "Attempting to open Excel file : $excelFile\n\n";

# get already active Excel application or open new
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
    || Win32::OLE->new('Excel.Application', 'Quit');  

# open Excel file
my $Book = $Excel->Workbooks->Open($excelFile) or die "Error! can't open excel file $excelFile. Make sure your file name and path are correct. Make sure you have access to the file.";

$theCount = $Book->Worksheets->Count;

# Give error and quit if excel book doesn't have 4 worksheets
if ($theCount > 4) {
  print "Warning! The number of worksheets is greater than four. This program is detecting $theCount worksheets. Only the first four worksheets will be processed.\n\n";
}

if ($theCount < 4) {
  print "Error! The number of worksheets is less than four. This program is detecting $theCount worksheets. Please make sure that there are four worksheets (even if they are blank) and re-run the program.";
  die;
}

# Header will be worksheet 1 
my $Header = $Book->Worksheets(1);
my $headerName = $Header->Name;

#Calculate last row for header
my $LastRowHeader = $Header->UsedRange->Find({What=>"*",
  SearchDirection=>xlPrevious,
  SearchOrder=>xlByRows})->{Row};

my $LastColHeader = $Header->UsedRange->Find({What=>"*", 
  SearchDirection=>xlPrevious,
  SearchOrder=>xlByColumns})->{Column};

if ($LastColHeader > 3) {
  print "Warning! worksheet 1, called $headerName has greater than three columns. This program detects $LastColHeader columns. This sheet should be the header for the output file and have three columns. Only the first three columns will be processed.\n\n";
}

if ($LastColHeader < 3) {
  print "Error! worksheet 1, called $headerName has less than three columns. This program detects $LastColHeader columns. This sheet should be the header for the output file and have three columns. Please modify the excel spreadsheet to have three columns (even if they are blank) and re-run this program.";
  die;
}

open(OUTPUTFILE, ">$outfile") or die "Error! can't open output file $outfile. Make sure you have access to the file."; 

$totalCount = 0;
$counter = 0;
print "Processing header worksheet named $headerName. Total number of rows = $LastRowHeader. Total number of columns = $LastColHeader. ";

#First, get header information and print it to output file, if header exists
foreach my $row(1..$LastRowHeader) {
  # skip empty cells
  next unless defined $Header->Cells($row,2)->{'Value'};

  # print . every 15 rows as it processes
  if ($row%30 == 0) {
    print ". ";
  }

  # grab value for current row and column
  my $cellName = $Header->Cells($row,2)->{'Value'}; 
  chomp($cellName);
  
  # write name to output file
  print OUTPUTFILE "$cellName";
  
  # if there is a value in column 3, output it to file
  if ($Header->Cells($row,3)->{'Value'}) {
    my $cellValue = $Header->Cells($row,3)->{'Value'}; 
    chomp($cellValue);
	my $resultValue = convertIfDate($cellValue);
    print OUTPUTFILE "$resultValue";
  }
  print OUTPUTFILE "\n";
  $counter++;
}
print "Finished processing the header worksheet named $headerName. $counter cells were processed.\n\n"; 
$totalCount += $counter;
print OUTPUTFILE "\n\n";

# process the 3 sheets
foreach my $sheetNumber(2..4) {
  my $Sheet = $Book->Worksheets($sheetNumber);
  processSheet($Sheet);
  print OUTPUTFILE "\n\n";
}

print "Finished processing all worksheets in excel sheet and outputting to file. $totalCount cells were processed total. The output is at : $outfile";

sub processSheet {
  my ($Sheet) = @_;
  my $sheetName = $Sheet->Name;
  my @FieldNames = ();

  # calculate last row and last column numbers
  my $LastRow = $Sheet->UsedRange->Find({What=>"*",
    SearchDirection=>xlPrevious,
    SearchOrder=>xlByRows})->{Row};

  my $LastCol = $Sheet->UsedRange->Find({What=>"*", 
    SearchDirection=>xlPrevious,
    SearchOrder=>xlByColumns})->{Column};

  #get column names of sheet (the second row)
  foreach my $col(1..$LastCol) {
    if (not defined $Sheet->Cells(2,$col)->{'Value'})
    {
      print "Warning! Undefined value in row #2, column #$col in the worksheet called $sheetName. (row #2 should have a field name for each column in it). This program will skip this column and all columns following it in this worksheet.\n\n";
      $newLastCol = $col - 1;
      last;
    }
    
    my $thisexcelline=$Sheet->Cells(2,$col)->{'Value'};
    chomp($thisexcelline);
   
    # Store column names
    push (@FieldNames,$thisexcelline);

    $totalCount++;
  }	

  #Don't need to process all the extra columns if there is no field name
  if ($newLastCol) {
    $LastCol = $newLastCol;
  }
  
  print "Processing the worksheet named $sheetName. Total number of rows = $LastRow. Total number of columns = $LastCol. ";
  $counter = 0;

  #process the sheet
  foreach my $row(3..$LastRow) {
    # print . every 15 rows as it processes
    if ($row%30 == 0) {
      print ". ";
    }
    
    foreach my $col(1..$LastCol) 
    {
      # skip empty cells
      next unless defined $Sheet->Cells($row,$col)->{'Value'};

      # grab value for current row and column
      my $thisexcelline=$Sheet->Cells($row,$col)->{'Value'};
      chomp($thisexcelline);
	  
	  my $result=convertIfDate($thisexcelline);
	  
      # write to output file
      print OUTPUTFILE "$FieldNames[$col - 1]$result\n";        
  
      # increment count
      $counter++;
    }
    

    print OUTPUTFILE "\n";
  }

  $totalCount += $counter;
  print "Finished processing the worksheet named $sheetName. $counter cells were processed.\n\n"; 
}

sub convertIfDate {
  my ($value) = @_;
  if( ref($value) eq 'Win32::OLE::Variant' and $value->Type == VT_DATE) {
    return $value->Date( 'MM-dd-yyyy' );
  }
  return $value;
}