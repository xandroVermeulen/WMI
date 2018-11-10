#init
use Win32::OLE qw(in); 
use Win32::OLE::Const;
$Win32::OLE::Warn = 3; 
my $filename1 = 'punten.xls';
my $filename2 = 'punten2.xls';
$fso = Win32::OLE->new("Scripting.FileSystemObject");
$excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
$excel->{DisplayAlerts}=0;
$excel->{visible} = 1; #open files
$book1 = open_excel_file($filename1);
$book2 = open_excel_file($filename2);
#sort file 2
my $student_amount = 189;
my $sheet_file2 = $book2->{Worksheets}->Item("perl");
$sheet_file2->{UsedRange}->Sort($sheet_file2->Range("A1:A".$student_amount));
#add file 2 to file 1
my $sheet_file1 = $book1->{Worksheets}->Item("architectuur");
$mat1=$sheet_file1->Range("B1:B".$student_amount)->{Value};$mat2=$sheet_file2->Range("B1:B".$student_amount)->{Value};
$sheet_file1->Range($sheet_file1->Cells(1,3),$sheet_file1->Cells(189,3))->{Value}=$mat2;
#Add column D
for ($row = 0 ; $row < $student_amount ; $row++){
	$avg=(${${$mat1}[$row]}[0] + ${${$mat2}[$row]}[0])/2;
	$mat_avg->[$row][0] = int($avg + 0.5);		
}
$sheet_file1->Range($sheet_file1->Cells(1,4),$sheet_file1->Cells(189,4))->{Value}=$mat_avg;
#nieuw worksheet maken
my $sheet_advalvas;
if($book1->{Worksheets}->{Count} == 1){
	$sheet_advalvas = $book1->{Worksheets}->Add();
	$sheet_advalvas->{Name} = "ad valvas";
} else {
	$sheet_advalvas = $book1->{Worksheets}->Item("ad valvas");	
}
#studenten in blokken verdelen
$mat_names = $sheet_file1->Range("A1:A".$student_amount)->{Value};
my @group_12=();
my @group_10=();
my @group_7=();
my @group_0=();

for ($row = 0 ; $row < $student_amount ; $row++){
	if(${${$mat_avg}[$row]}[0] > 11){
		push @group_12, ${${$mat_names}[$row]}[0];
	} elsif(${${$mat_avg}[$row]}[0] >9){
		push @group_10, ${${$mat_names}[$row]}[0];
	} elsif(${${$mat_avg}[$row]}[0] >6){
		push @group_7, ${${$mat_names}[$row]}[0];
	} else {
		push @group_0, ${${$mat_names}[$row]}[0];
	}
}
my $next_row = 1;
$next_row  = create_block('A',5,$next_row,\@group_12);$next_row = create_block('B',5,$next_row,\@group_10);
$next_row = create_block('C',5,$next_row,\@group_7);
$next_row = create_block('D',5,$next_row,\@group_0);

$sheet_advalvas->{UsedRange}->{Columns}->AutoFit();
$sheet_advalvas->{UsedRange}->{Interior}->{ColorIndex} = 2;save();
print "Press ENTER to exit"; <STDIN>;

sub create_block{
	my ($name,$column_count,$start_row,$array_ref) = splice @_,0,4;
	my $student_count = @$array_ref;
	my $max_rows = ($student_count-($student_count%($column_count-1)))/($column_count-1);	
	$cell = $sheet_advalvas->Cells($start_row+1,($column_count - ($column_count%2)) / 2  + 1);
	$cell->{Value}=$name;
	$cell->{Font}->{Bold} = 1;	
	
	for ($i = 0 ; $i < $student_count ; $i++){
		$remainder = $i%$max_rows;
		$mat->[$remainder][($i - $remainder) / $max_rows] = ${$array_ref}[$i];
	}
	my $end_row = $start_row+1+$max_rows;
	$group_range =$sheet_advalvas->Range($sheet_advalvas->Cells($start_row+2,1),$sheet_advalvas->Cells($end_row, $column_count));
	$group_range->{Value}=$mat;
	$group_range->BorderAround(6);	return $end_row+1;
}sub open_excel_file{
	my $filename = shift;
	my $book;
	if ($fso->FileExists($filename)) {
		my $path = $fso->GetAbsolutePathName($filename);
		$book=$excel->{Workbooks}->Open($path); 
	}
	else {
	    my $dir = $fso->GetAbsolutePathName("."); 
	    my $path = $dir."\\".$filename;            
	    $book = $excel->{Workbooks}->Add();   
	    $book->SaveAs($path);                 
	}
	return $book;
}
sub save{
	$book1->Save();
}