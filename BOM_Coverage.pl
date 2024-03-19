#!/usr/bin/perl
print "\n";
print "*******************************************************************************\n";
print "  Bom Coverage ckecking tool for 3070 <v5.91>\n";
print "  Author: Noon Chen\n";
print "  A Professional Tool for Test.\n";
print "  ",scalar localtime;
print "\n*******************************************************************************\n";
print "\n";

#v5.90	152s
#v5.91	143s handled components
#v5.91	126s handled untestable
#v5.91	93s  handled parallel

#########################################################################################
#print "  Checking In: ";
use Term::ReadKey;
use Time::HiRes qw(time);

# system "stty -echo";
 ReadMode('noecho'); # Disable echoing of characters
 print "  Password: ";
 chomp($Ccode = <STDIN>);
 print "\n";
# system "stty echo";
 ReadMode('restore'); # Restore terminal mode

   if ($Ccode ne "\@testpro")
   #if ($Ccode ne "TestPro")
    	{
    		print "  >>> Wrong password!\n"; goto END_Prog;
    	}
   else
   	{
    		print "  >>> Correct password.\n\n";
   	}


############################ Excel ######################################################
use Excel::Writer::XLSX;

my $bom_coverage_report = Excel::Writer::XLSX->new('BOM_Coverage_Report.xlsx');
my $summary = $bom_coverage_report-> add_worksheet('Summary');
my $coverage = $bom_coverage_report-> add_worksheet('Coverage');
my $tested = $bom_coverage_report-> add_worksheet('Tested');
my $untest = $bom_coverage_report-> add_worksheet('Untest');
my $limited = $bom_coverage_report-> add_worksheet('LimitTest');
my $power = $bom_coverage_report-> add_worksheet('PowerTest');
my $short_thres = $bom_coverage_report-> add_worksheet('Shorts_Thres');

$coverage-> freeze_panes(1,1);			#冻结行、列
$tested-> freeze_panes(1,1);			#冻结行、列
$untest-> freeze_panes(1,0);			#冻结行、列
$limited-> freeze_panes(1,0);			#冻结行、列
$power-> freeze_panes(1,0);				#冻结行、列
$short_thres-> freeze_panes(1,0);		#冻结行、列

$summary-> set_column(0,2,20);			#设置列宽
$coverage-> set_column('A:F',20);		#设置列宽
$tested-> set_column('A:E',20);			#设置列宽
$tested-> set_column('F:F',40);			#设置列宽
$untest-> set_column(0,1,20);			#设置列宽
$untest-> set_column(1,2,40);			#设置列宽
$limited-> set_column(0,1,20);			#设置列宽
$limited-> set_column(1,1,30);			#设置列宽
$power-> set_column(0,1,20);			#设置列宽
$power-> set_column(1,3,30);			#设置列宽
$power-> set_column(4,12,15);			#设置列宽
$short_thres-> set_column(0,1,40);		#设置列宽

$summary-> activate();					#设置初始可见

#新建一个格式
$format_item = $bom_coverage_report-> add_format(bold=>1, align=>'center', valign=>'vcenter', border=>1, size=>12, bg_color=>'cyan');
$format_head = $bom_coverage_report-> add_format(bold=>1, valign=>'vcenter', border=>1, size=>12, bg_color=>'lime');
$format_data = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, text_wrap=>1);
$format_GND  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'gray');
$format_NC   = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'silver');
$format_VCC  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'orange');
$format_togg = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'green', text_wrap=>1);
$format_pin  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'lime', text_wrap=>1);
$format_anno = $bom_coverage_report-> add_format(align=>'left', valign=>'vcenter', border=>1, text_wrap=>1);
$format_anno1 = $bom_coverage_report-> add_format(align=>'left', valign=>'vcenter', border=>1, text_wrap=>1, bg_color=>'yellow');
$format_PCT  = $bom_coverage_report-> add_format(align=>'center', border=>1, num_format=> '10');
$format_STP  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'yellow');
$format_hylk = $bom_coverage_report-> add_format(color=>'blue', align=>'center', valign=>'vcenter', border=>1, underline=>1);
$format_FPY  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, num_format=> '10');

$row = 0; $col = 0;	$tested-> write($row, $col, '<Device>', $format_head);
$row = 0; $col = 1;	$tested-> write($row, $col, '<TYPE>', $format_head);
$row = 0; $col = 2;	$tested-> write($row, $col, '<Nominal>', $format_head);
$row = 0; $col = 3;	$tested-> write($row, $col, '<HiLimit>', $format_head);
$row = 0; $col = 4;	$tested-> write($row, $col, '<LoLimit>', $format_head);
$row = 0; $col = 5;	$tested-> write($row, $col, '<Comments>', $format_head);

$row = 0; $col = 0;	$untest-> write($row, $col, '<Device>', $format_head);
$row = 0; $col = 1;	$untest-> write($row, $col, '<Justification>', $format_head);
$row = 0; $col = 2;	$untest-> write($row, $col, '<Comments>', $format_head);

$row = 0; $col = 0;	$limited-> write($row, $col, '<Device>', $format_head);
$row = 0; $col = 1;	$limited-> write($row, $col, '<Comments>', $format_head);

$row = 0; $col = 0;	$power-> write($row, $col, '<Device>', $format_head);
$row = 0; $col = 1;	$power-> write($row, $col, '<TestOrder>', $format_head);
$row = 0; $col = 2;	$power-> write($row, $col, '<TestPlan>', $format_head);
$row = 0; $col = 3;	$power-> write($row, $col, '<Family>', $format_head);
$row = 0; $col = 4;	$power-> write($row, $col, '<Total Pin>', $format_item);
$row = 0; $col = 5;	$power-> write($row, $col, '<Power Pin>', $format_VCC);
$row = 0; $col = 6;	$power-> write($row, $col, '<GND Pin>', $format_GND);
$row = 0; $col = 7;	$power-> write($row, $col, '<Toggle Test Pin>', $format_togg);
$row = 0; $col = 8;	$power-> write($row, $col, '<Pin Test>', $format_pin);
$row = 0; $col = 9;	$power-> write($row, $col, '<NC Pin>', $format_NC);
$row = 0; $col = 10;	$power-> write($row, $col, '<Untest Pin>', $format_data);
$row = 0; $col = 11;	$power-> write($row, $col, '<Toggle Coverage>', $format_togg);
$row = 0; $col = 12;	$power-> write($row, $col, '<Pin Coverage>', $format_pin);

$row = 0; $col = 0;	$short_thres-> write($row, $col, 'Nodes', $format_head);
$row = 0; $col = 1;	$short_thres-> write($row, $col, 'Threshold', $format_head);

$row = 0; $col = 0;	$summary-> write($row, $col, 'Test Items', $format_head);
$row = 0; $col = 1;	$summary-> write($row, $col, 'Quantity', $format_head);
$row = 0; $col = 2;	$summary-> write($row, $col, 'Percentage', $format_head);

$row = 1; $col = 0;	$summary-> write($row, $col, 'Tested', $format_item);
$row = 2; $col = 0;	$summary-> write($row, $col, 'Untest', $format_item);
$row = 3; $col = 0;	$summary-> write($row, $col, 'LimitTest', $format_item);
$row = 4; $col = 0;	$summary-> write($row, $col, 'Power-Tested', $format_item);
$row = 5; $col = 0;	$summary-> write($row, $col, 'Power-UnTest', $format_item);
$row = 6; $col = 0;	$summary-> write($row, $col, 'Node accessibility rate', $format_item);

$row = 1; $col = 1;	$summary-> write($row, $col, '=COUNTA(Tested!A2:A9999)', $format_data);
$row = 2; $col = 1;	$summary-> write($row, $col, '=COUNTA(Untest!A2:A9999)', $format_data);
$row = 3; $col = 1;	$summary-> write($row, $col, '=COUNTA(LimitTest!A2:A9999)', $format_data);
$row = 4; $col = 1;	$summary-> write($row, $col, '=COUNTA(PowerTest!A2:A9999)-B6', $format_data);
$row = 6; $col = 1;
$summary-> write($row, $col, '=COUNTA(Shorts_Thres!A2:A9999)-COUNTIF(Shorts_Thres!A2:A9999,"!nodes *")', $format_data);

$row = 1; $col = 2;	$summary-> write_formula($row, $col, "=(B2/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$row = 2; $col = 2;	$summary-> write_formula($row, $col, "=(B3/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$row = 3; $col = 2;	$summary-> write_formula($row, $col, "=(B4/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$row = 4; $col = 2;	$summary-> write_formula($row, $col, "=(B5/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$row = 5; $col = 2;	$summary-> write_formula($row, $col, "=(B6/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$row = 6; $col = 2;	$summary-> write_formula($row, $col, "=(B7/COUNTA(Shorts_Thres!A2:A9999))", $format_PCT);  #输出Percentage

$tested-> write("H2", 'Type', $format_item);
$tested-> write("H3", 'Count', $format_item);
$tested-> write("I2", 'Resistor', $format_head);
$tested-> write("J2", 'Capacitor', $format_head);
$tested-> write("K2", 'Inductor', $format_head);
$tested-> write("L2", 'Jumper', $format_head);
$tested-> write("M2", 'Diode', $format_head);
$tested-> write("N2", 'Zener', $format_head);
$tested-> write_formula("I3", '=COUNTIF(B2:B99999,"Resistor")', $format_data);
$tested-> write_formula("J3", '=COUNTIF(B2:B99999,"Capacitor")', $format_data);
$tested-> write_formula("K3", '=COUNTIF(B2:B99999,"Inductor")', $format_data);
$tested-> write_formula("L3", '=COUNTIF(B2:B99999,"Jumper")', $format_data);
$tested-> write_formula("M3", '=COUNTIF(B2:B99999,"Diode")', $format_data);
$tested-> write_formula("N3", '=COUNTIF(B2:B99999,"Zener")', $format_data);

$row = 0; $col = 0;	$coverage-> write($row, $col, ' Test Items', $format_head);
$row = 0; $col = 1;	$coverage-> write($row, $col, ' L, C, R, D, Z, J test', $format_head);
$row = 0; $col = 2;	$coverage-> write($row, $col, ' Digital logic test', $format_head);
$row = 0; $col = 3;	$coverage-> write($row, $col, ' Analog function test', $format_head);
$row = 0; $col = 4;	$coverage-> write($row, $col, ' Bscan test', $format_head);
$row = 0; $col = 5;	$coverage-> write($row, $col, ' No coverage', $format_head);

$power-> conditional_formatting('H2:H9999',
    {
    	type     => 'cell',
     	criteria => 'greater than',
     	value    => 0,
     	format   => $format_togg,
    });
$power-> conditional_formatting('L2:L9999',
    {
    	type     => 'cell',
     	criteria => 'greater than',
     	value    => 0,
     	format   => $format_togg,
    });

$power-> conditional_formatting('I2:I9999',
    {
    	type     => 'cell',
     	criteria => 'greater than',
     	value    => 0,
     	format   => $format_pin,
    });
$power-> conditional_formatting('M2:M9999',
    {
    	type     => 'cell',
     	criteria => 'greater than',
     	value    => 0,
     	format   => $format_pin,
    });

my $chart = $bom_coverage_report-> add_chart( type => 'pie', embedded => 1 );
$chart-> add_series(
    name       => '=Summary!$C$1',
    categories => '=Summary!$A$2:$A$6',
    values     => '=Summary!$B$2:$B$6',
    data_labels => {value => 1},
	);
$chart-> set_style( 10 );
$summary-> insert_chart('A10', $chart, 10, 0, 1.0, 1.6);

$rowC = 0;
$rowT = 1;
$rowU = 1;
$rowL = 1;
$rowP = 1;
$length_anno = 8;
$length_TO = 8;
$length_TP = 8;
$PowerUT = 0;

print "  please specify BOM list file: ";
   $bom=<STDIN>;
   chomp $bom;

$start_time = time();

open (BOM, "< $bom");                      #read BOM list.

print "\n";
print "  gether all BOM devices...";
@bom_list = <BOM>;
print "[DONE]\n";

my %hash_pin = ();
open (Pin, "< pins") || open (Pin, "< 1%pins"); 
	while($nodes = <Pin>)
	{
		$nodes =~ s/(^\s+|\s+$)//g;
		@nodes = split('\"',$nodes);
		if ($nodes[1] =~ "\%"){$nodes[1] = substr($nodes[1],2);}
		
		if (substr($nodes[0], 0, 5) eq "nodes")
		{
		$hash_pin{$nodes[1]} = 1;
		#print $nodes[0],"-- ",$nodes[1]."\n";
		}
	}
close Pin;
#$count = 0;

foreach $device (@bom_list)
{
	$Total_Pin = 0;
	$Power_Pin = 0;
	$GND_Pin = 0;
	$Toggle_Pin = 0;
	$NC_Pin = 0;
	$Untest_Pin = 0;
	#-------------------------------------------------------------------------------------
	$worksheet = 0;
	$foundTO = 0;
	$foundTP = 0;
	$device =~ s/(^\s+|\s+$)//g;                     #clear all spacing
	$device = lc($device);

	$rowC = $rowC+1;	$cover = 0;	$UNCover = 0;
	$coverage-> write($rowC, 0, $device, $format_data);			#Coverage
	$coverage-> write($rowC, 1, "-", $format_data);
	$coverage-> write($rowC, 2, "-", $format_data);
	$coverage-> write($rowC, 3, "-", $format_data);
	$coverage-> write($rowC, 4, "-", $format_data);
	$coverage-> write($rowC, 5, "-", $format_data);
	$len_device = length($device);
	print "	Analyzing ", $device, " .....\n";
	#************************ Testorder Checking *******************************************
	open (TesO, "<testorder");																#### check testorder ####
	$len = 0;
		while($lineTO = <TesO>)
			{
			#$count = $count+1;
			#print $count."\n";
			chomp($lineTO);
			$lineTO =~ s/(^\s+|\s+$)//g; 					 #clear all spacing
			@DevTO =  split('\"', $lineTO);
			$DevTO[1] =~ s/(^\s+|\s+$)//g;
					$nullTO = index($lineTO,"nulltest");
					$skipTO = index($lineTO,"skip");
					$comment1 = index($lineTO,"\;");
					$comment2 = index($lineTO,"\!");
					$scan = index($lineTO,"scan");
					$powered = index($lineTO,"analog powered");
					$digital = index($lineTO,"digital");
					$mixed = index($lineTO,"mixed");
					$ver = index($lineTO,"version");
			################ testable device ##############################################################################################
			if($DevTO[1] eq $device		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\%")		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\_")
				and substr($DevTO[1],0,length($device)) eq $device
          and $nullTO == -1
          and $skipTO == -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					$learn = 0;
					print "			General Test ", $DevTO[1],"\n";   #, $lineTO,"\n";
			$tested-> write($rowT, 2, "-", $format_data);
			$tested-> write($rowT, 3, "-", $format_data);
			$tested-> write($rowT, 4, "-", $format_data);

					$testname = $DevTO[1];
				  	#print $testname,"\n";	print length($testname),"\n";
					open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/(^\s+|\s+$)//g;									#clear head of line spacing
							if ($lineTP =~ "learn capacitance on"){$learn = 1;}
							if ($lineTP =~ "learn capacitance off"){$learn = 0;}
							next if ($learn == 1);
							@DevTP =  split('\"', $lineTP);
							$DevTP[1] =~ s/(^\s+|\s+$)//g;

							if($DevTP[1] eq $testname	|substr($DevTP[1],7) eq $testname	#matching test name
								and substr($DevTP[1],0,length($device)) eq $device	|substr($DevTP[1],7,length($device)) eq $device
								and substr($lineTP,0,4) eq "test")							#matching not skipped test name
								{
									$foundTP = 1;
									$UNCover = 1;
								#print $lineTP,"\n"; print substr($lineTP,0,1)."\n";
								if (substr($DevTP[1],7) eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = $DevTP[1];}
								if ($DevTP[1] eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = "analog/".$testname;}
								if (index($lineTP,"on boards")> -1)
									{$testfile = "analog/1%".$testname;}

								if ($testfile eq $testfile_last){last;}						#ignore duplicated test name
  									$testfile_last = $testfile;
									$commentTP = "";
								if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"));}   #TP comments
  									$coverage-> write($rowC, 1, 'V', $format_togg);			#Coverage

								open (SourceFile, "<$testfile");
									while($lineTF = <SourceFile>)							#read parameter
									{
										$len = 0;
										chomp;
										$lineTF =~ s/^ +//;                               	#clear head of line spacing
										#print $lineTF;
										if (substr($lineTF,0,8) eq "resistor")				#### matching resistor ########
										{
											$tested-> write($rowT, 0, $testname, $format_item);  						## TestName ##
											$tested-> write($rowT, 1, substr($lineTF,0,8), $format_data);  				## TestType ##
											#print substr($lineTF,9)."\n";
											@param =  split('\,', substr($lineTF,9));
											$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
											if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
											if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
											if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
											if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
											$tested-> write($rowT, 5, $commentTP, $format_anno);  						## Comment ##
											#@array = ($testname,substr($lineTF,0,8),$param[0],$param[1],$param[2],$commentTP);
											#$array_ref = \@array;
											#$tested-> write_row($rowT, 0, $array_ref, $format_data);
											$rowT++;
											if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
											last; #goto Next_Dev;
										}
										elsif (substr($lineTF,0,9) eq "capacitor")				#### matching capacitor ########
										{
											$tested-> write($rowT, 0, $testname, $format_item);  						## TestName ##
											$tested-> write($rowT, 1, substr($lineTF,0,9), $format_data);  				## TestType ##
											#print substr($lineTF,9)."\n";
											@param =  split('\,', substr($lineTF,10));
											$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
											if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
											if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
											if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
											if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
											$tested-> write($rowT, 5, $commentTP, $format_anno);  						## Comment ##
											#@array = ($testname,substr($lineTF,0,9),$param[0],$param[1],$param[2],$commentTP);
											#$array_ref = \@array;
											#$tested-> write_row($rowT, 0, $array_ref, $format_data);
											$rowT++;
											if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
											last; #goto Next_Dev;
										}
										elsif (substr($lineTF,0,8) eq "inductor")					#### matching inductor ########
										{
											$tested-> write($rowT, 0, $testname, $format_item);  						## TestName ##
											$tested-> write($rowT, 1, substr($lineTF,0,8), $format_data);  				## TestType ##
											#print substr($lineTF,9)."\n";
											@param =  split('\,', substr($lineTF,9));
											$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
											if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
											if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
											if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
											if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
											$tested-> write($rowT, 5, $commentTP, $format_anno);  						## Comment ##
											$rowT++;
											if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
											last; #goto Next_Dev;
										}
										elsif (substr($lineTF,0,5) eq "diode")					#### matching diode ######
										{
											$tested-> write($rowT, 0, $testname, $format_item);  						## TestName ##
											$tested-> write($rowT, 1, substr($lineTF,0,5), $format_data);  				## TestType ##
											#print substr($lineTF,9)."\n";
											@param =  split('\,', substr($lineTF,6));
											if ($lineTF !~ m/\"/g){
											#$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
											$tested-> write($rowT, 3, $param[0], $format_data);  						## HiLimit ##
											$tested-> write($rowT, 4, $param[1], $format_data);	  						## LoLimit ##
											$tested-> write($rowT, 5, $commentTP, $format_anno);  						## Comment ##
											}
											if ($lineTF =~ m/\"/g){
											$DioNom = "";	$DioNom =  $DioNom . $param[0];  							## Nominal ## 
											$DioHiL = "";	$DioHiL =  $DioHiL . $param[1];  							## HiLimit ##
											$DioLoL = "";	$DioLoL =  $DioLoL . $param[2];	  							## LoLimit ##
											$tested-> write($rowT, 5, $commentTP, $format_anno);  						## Comment ##
											while($lineTF = <SourceFile>){
												chomp;
												$lineTF =~ s/^ +//;
												if (substr($lineTF,0,5) eq "diode"){
												#print $lineTF."\n";
												@param =  split('\,', substr($lineTF,6));
													$DioNom =  $DioNom . "\n" . $param[0];  							## Nominal ## 
													$DioHiL =  $DioHiL . "\n" . $param[1];  							## HiLimit ##
													$DioLoL =  $DioLoL . "\n" . $param[2];	  							## LoLimit ##
													}
												elsif (eof){
													$tested-> write($rowT, 2, $DioNom, $format_data);  					## Nominal ## 
													$tested-> write($rowT, 3, $DioHiL, $format_data);  					## HiLimit ##
													$tested-> write($rowT, 4, $DioLoL, $format_data);	  				## LoLimit ##
													last;}
												}
											}
											$rowT++;
											last; #goto Next_Dev;
										}
										elsif (substr($lineTF,0,5) eq "zener")					 ####matching zener##
										{
											$tested-> write($rowT, 0, $testname, $format_item);  						## TestName ##
											$tested-> write($rowT, 1, substr($lineTF,0,5), $format_data);  				## TestType ##
											#print substr($lineTF,9)."\n";
											@param =  split('\,', substr($lineTF,6));
											$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
											if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
											if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
											if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
											if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
											$tested-> write($rowT, 5, $commentTP, $format_anno);  						## Comment ##
											$rowT++;
											if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
											last; #goto Next_Dev;
										}
										elsif (substr($lineTF,0,6) eq "jumper")						####matching jumper########
										{
											$tested-> write($rowT, 0, $testname, $format_item);  						## TestName ##
											$tested-> write($rowT, 1, substr($lineTF,0,6), $format_data);  				## TestType ##
											#print substr($lineTF,9)."\n";
											@param =  split('\,', substr($lineTF,6));
											$OP = 0; if ($lineTF =~ "op") {$OP = 1;}
											if($OP == 0){$tested-> write($rowT, 3, $param[0], $format_data);	$tested-> write($rowT, 5, $commentTP, $format_anno);}  ## Excel ##
											if($OP == 1){$tested-> write($rowT, 4, $param[0], $format_STP);	$tested-> write($rowT, 5, "OP test", $format_STP);}	## Excel ##
											$rowT++;
											if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
											last; #goto Next_Dev;
										}
										elsif (substr($lineTF,0,4) eq "fuse")					  	####matching fuse########
										{
											$tested-> write($rowT, 0, $testname, $format_item);  						## TestName ##
											$tested-> write($rowT, 1, substr($lineTF,0,6), $format_data);  				## TestType ##
											#print substr($lineTF,9)."\n";
											@param =  split('\,', substr($lineTF,6));
											$tested-> write($rowT, 3, $param[0], $format_data);
											$tested-> write($rowT, 5, $commentTP, $format_anno);  ## Excel ##
											$rowT++;
											if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
											last; #goto Next_Dev;
										}
										elsif (eof)					  														####no parameter######
										{
										$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
										$untest-> write($rowU, 1, "No Test Parameter Found in TestFile.", $format_data);  	## Excel ##
										$rowU++;
										if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
										last;
										}
  								}
								}
							elsif ($DevTP[1] eq $testname	|substr($DevTP[1],7) eq $testname		#matching skipped test in TP
								and substr($DevTP[1],0,length($device)) eq $device	|substr($DevTP[1],7,length($device)) eq $device
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}																		#ignore duplicated test name
  									$testname_last = $testname;
									if ($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_NC);}			#Coverage
									$untest-> write($rowU, 0, $testname, $format_data);		## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
									if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
									last;
								}
							elsif (eof and $foundTP == 0){
							$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
							$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							$untest-> write($rowU, 2, $lineTO, $format_anno);  ## Excel ##
							$rowU++;
							goto Next_Dev;}
							}
					close TesP;
				}
			################ untestable devices ###########################################################################################
			elsif($DevTO[1] eq $device		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\%")		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\_")
			and substr($DevTO[1],0,length($device)) eq $device
          and $nullTO > -1
          and $comment1 > -1
          and $comment2 == -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					print "			NULL_Test  ", $DevTO[1],"\n";   #, $lineTO,"\n";

					$untest-> write($rowU, 0, $DevTO[1], $format_data);  ## Excel ##
					$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_data);  ## Excel ##
				$UTline = "";

				if($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_NC);}			#Coverage
				open(ALL, "<analog/$DevTO[1]")||open(ALL, "<analog/1%$DevTO[1]") or $untest-> write($rowU, 2, "!TestFile not found.", $format_anno);
				while($line = <ALL>)
					{
					if (index($line,$device)>1){
						$line = substr($line,1);
						$line =~ s/(^\s+)//g;
						if (length($line)> $length_anno){$length_anno = length($line);}
						$UTline = $line . $UTline;}
					elsif (eof){last;}
					}
					$UTline =~ s/(^\s+|\s+$)//g;
					$untest-> write($rowU, 2, $UTline, $format_anno);
					$untest-> set_column(2, 2, $length_anno);
				$rowU++;
				close ALL;
				if (($DevTO[1] =~ $device) and ($DevTO[0] =~ "capacitor|resistor|jumper") and (substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
				# last;
				}
			################ parallel tested devices ######################################################################################
			elsif($DevTO[1] eq $device		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\%")		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\_")
				and substr($DevTO[1],0,length($device)) eq $device
          and $nullTO > -1
          and $comment1 > -1
          and $comment2 > -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					print "			Parallel_Test  ", $DevTO[1],"\n";		#, $lineTO,"\n";
					$anno = substr($lineTO,rindex($lineTO,"\!")+1);
					$anno =~ s/(^\s+|\s+$)//g; 
					$coverage-> write($rowC, 1, 'L', $format_pin);			#Coverage
					$limited-> write($rowL, 0, $DevTO[1], $format_data);	## Excel ##
					$limited-> write($rowL, 1, $anno, $format_anno);		## Excel ##
					$rowL++;
					if (($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\%") and ($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) ne "\_")){goto Next_Dev;}
					# last;
				}
			################ testable analog powered test #################################################################################
			elsif($DevTO[1] eq $device		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\%")		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\_")
				and substr($DevTO[1],0,length($device)) eq $device
          and $powered > -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					#$cover = 0;
					print "			ANA_PWD_Test  ", $DevTO[1],"\n";   #, $lineTO,"\n";
			$power-> write($rowP, 4, "-", $format_data);
			$power-> write($rowP, 5, "-", $format_data);
			$power-> write($rowP, 6, "-", $format_data);
			$power-> write($rowP, 7, "-", $format_data);
			$power-> write($rowP, 8, "-", $format_data);
			$power-> write($rowP, 9, "-", $format_data);
			$power-> write($rowP, 10, "-", $format_data);
			$power-> write($rowP, 11, "-", $format_data);
			$power-> write($rowP, 12, "-", $format_data);

				 		$testname = $DevTO[1];
						open (TesP, "<testplan");											#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/(^\s+|\s+$)//g;									#clear head of line spacing
							@DevTP =  split('\"', $lineTP);
							$DevTP[1] =~ s/(^\s+|\s+$)//g;

							if($DevTP[1] eq $testname	|substr($DevTP[1],7) eq $testname	#matching test name
								and substr($DevTP[1],0,length($device)) eq $device	|substr($DevTP[1],7,length($device)) eq $device)
							# and substr($lineTP,0,4) eq "test")							#matching not skipped test name
								{
									$foundTP = 1;
								#print $lineTP,"\n"; print substr($lineTP,0,1)."\n";
								if (substr($DevTP[1],7) eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = $DevTP[1];}
								if ($DevTP[1] eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = "analog/".$testname;}
								if (index($lineTP,"on boards")> -1)
									{$testfile = "analog/1%".$testname;}

								#print $device,"\n";
								if ($testfile eq $testfile_last){last;}					#ignore duplicated test name
  									$testfile_last = $testfile;
									$commentTP = "";
								if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"));}   #TP comments
								
									$power-> write($rowP, 0, $device, $format_data);		## Excel ##
									$power-> write($rowP, 1, $lineTO, $format_anno);		## Excel ##
									if (length($lineTO)	> $length_TO){$length_TO = length($lineTO); $power-> set_column(1, 1, $length_TO);}
									if(substr($lineTP,0,4) eq "test"){
									$cover = 1;
									$coverage-> write($rowC, 3, 'V', $format_togg);			#Coverage
									$power-> write($rowP, 2, $lineTP, $format_anno);		## Excel ##
									if (length($lineTP)	> $length_TP){$length_TP = length($lineTP); $power-> set_column(2, 2, $length_TP);}
									}
									if(substr($lineTP,0,1) eq "\!"){
									$PowerUT = $PowerUT + 1;
									if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}			#Coverage
									$power-> write($rowP, 2, $lineTP, $format_anno1);		## Excel ##
									if (length($lineTP)	> $length_TP){$length_TP = length($lineTP); $power-> set_column(2, 2, $length_TP);}
									}
									$rowP++;
									last;
  								}
							elsif (eof and $foundTP == 0){
							$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
							$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							$untest-> write($rowU, 2, $lineTO, $format_anno);  ## Excel ##
							$rowU++;
							goto Next_Dev;}
							}
						close TesP;
				}
			################ testable digital test ########################################################################################
			elsif($DevTO[1] eq $device		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\%")		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\_")
			and substr($DevTO[1],0,length($device)) eq $device
          and $powered == -1
          and $scan == -1
          and $digital > -1
          and $mixed == -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					#$cover = 0;
					$length_DigPin = 10;
					print "			Digital_Test  ", $DevTO[1],"\n";   #, $lineTO,"\n";
			$power-> write($rowP, 4, "-", $format_data);
			$power-> write($rowP, 5, "-", $format_data);
			$power-> write($rowP, 6, "-", $format_data);
			$power-> write($rowP, 7, "-", $format_data);
			$power-> write($rowP, 8, "-", $format_data);
			$power-> write($rowP, 9, "-", $format_data);
			$power-> write($rowP, 10, "-", $format_data);
			$power-> write($rowP, 11, "-", $format_data);
			$power-> write($rowP, 12, "-", $format_data);

				 		$testname = $DevTO[1];
						open (TesP, "<testplan");											#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/(^\s+|\s+$)//g;									#clear head of line spacing
							@DevTP =  split('\"', $lineTP);
							$DevTP[1] =~ s/(^\s+|\s+$)//g;

							if($DevTP[1] eq $testname	|substr($DevTP[1],8) eq $testname	#matching test name
								and substr($DevTP[1],0,length($device)) eq $device	|substr($DevTP[1],8,length($device)) eq $device)
							# and substr($lineTP,0,4) eq "test")							#matching not skipped test name
								{
									$foundTP = 1;
								#print $lineTP,"\n"; print substr($DevTP[1],8,length($device))."\n";
								if (substr($DevTP[1],8) eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = $DevTP[1];}
								if ($DevTP[1] eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = "digital/".$testname;}
								if (index($lineTP,"on boards")> -1)
									{$testfile = "digital/1%".$testname;}

								if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
  									$testfile_last = $testfile;
									$commentTP = "";
								if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"));}   #TP comments

									$power-> write($rowP, 1, $lineTO, $format_anno);  			## Excel ##
									if (length($lineTO)	> $length_TO){$length_TO = length($lineTO); $power-> set_column(1, 1, $length_TO);}

									if(substr($lineTP,0,4) eq "test"){
									$cover = 1;
									$coverage-> write($rowC, 2, 'V', $format_togg);				#Coverage
									$power-> write($rowP, 2, $lineTP, $format_anno);			## Excel ##
									if (length($lineTP)	> $length_TP){$length_TP = length($lineTP); $power-> set_column(2, 2, $length_TP);}
									}
									if(substr($lineTP,0,1) eq "\!"){
									$PowerUT = $PowerUT + 1;
									if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_NC);}			#Coverage
									$power-> write($rowP, 2, $lineTP, $format_anno1);			## Excel ##
									if (length($lineTP)	> $length_TP){$length_TP = length($lineTP); $power-> set_column(2, 2, $length_TP);}
									}
									###### hyperlink #####################################
									$power-> write_url($rowP, 0, 'internal:'.$device.'!A1');	## hyperlink
									
									if ($worksheet == 0){
									$worksheet = 1;
									my $IC = $bom_coverage_report-> add_worksheet($device);		## hyperlink
									$IC-> write_url('A1', 'internal:PowerTest!A1');  			## hyperlink
									$IC->conditional_formatting('A1:GR999',
									    {
									    	type     => 'text',
									     	criteria => 'containing',
									     	value    => 'Toggle_Test',
									     	format   => $format_togg,
									    });
									
									$IC->conditional_formatting('A1:GR999',
									    {
									    	type     => 'text',
									     	criteria => 'containing',
									     	value    => 'Contact_Test',
									     	format   => $format_pin,
									    });
									
									if ($length_DigPin < 10){$length_DigPin = 10;}
									open (Boards, "< board");
									while($lineDig = <Boards>)								
									{
										$lineDig =~ s/(^\s+|\s+$)//g;
										#print $lineDig;
										if (substr($lineDig,0,7) eq "DEVICES")
											{
											while($lineDig = <Boards>)								
												{
												$lineDig =~ s/(^\s+|\s+$)//g;
												$testname1 = uc($testname);
												#print $testname1."\n";
												#print $lineDig."\n";
												if ($lineDig eq uc($device))
													{while($lineDig = <Boards>) 
														{#print $lineDig;
														$Total_Pin++;
														@DigPin = split('\.',$lineDig);
														$DigPin[0] =~ s/(^\s+|\s+$)//g;
														$DigPin[1] =~ s/(^\s+|\s+$)//g;
														#print $DigPin[0]."\n";
														if ($DigPin[1] =~ /(GND|GROUND)/){
															if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_GND); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
															if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1], $format_GND);
															($pos) = $DigPin[0] =~ /^\D+/g;
															if (length($pos) == 1){$DigPos = ord($pos)%64;}
															if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
															#print $DigPos."\n";
															if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
															}
															$GND_Pin++;
														}
														elsif ($DigPin[1] =~ /(^\+0|^0V|^\+1|^1V|^\+2|^2V|^\+3|^3V|^\+5|^5V|^V_|^VCC|^VDD|^PP|^P0V|^P1V|^P2V|^P3V|^P5V)/){
															if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_VCC); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
															if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1], $format_VCC);
															($pos) = $DigPin[0] =~ /^\D+/g;
															if (length($pos) == 1){$DigPos = ord($pos)%64;}
															if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
															#print $DigPos."\n";
															if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
															}
															$Power_Pin++;
														}
														elsif ($DigPin[1] =~ /(^NC_|_NC$|NONE)/){
															if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_NC); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
															if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1], $format_NC);
															($pos) = $DigPin[0] =~ /^\D+/g;
															if (length($pos) == 1){$DigPos = ord($pos)%64;}
															if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
															#print $DigPos."\n";
															if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
															}
															$NC_Pin++;
														}
														else{
														my $BDG_File = "./bdg_data/dig_inc_ver_fau.dat";
														if(-e $BDG_File){
															open (BDGFile, "< ./bdg_data/dig_inc_ver_fau.dat");
															while($BDGline = <BDGFile>)
															{
															$BDGline =~ s/(^\s+|\s+$)//g;
															@BDG = split('\"',$BDGline);
															if ($BDG[1] =~ "\%"){$BDG[1] = substr($BDG[1],2);}
															if ($BDG[3] =~ "\%"){$BDG[3] = substr($BDG[3],2);}
															@BDGDig = split('\.',$BDG[1]);
															#print $BDGline."\n";
															if($BDGDig[0] eq $device and $BDGDig[1] eq $DigPin[0] and $BDG[3] eq $DigPin[1])
																{
																if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Toggle_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
																if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1]."\n* Toggle_Test", $format_data);
																	($pos) = $DigPin[0] =~ /^\D+/g;
																	if (length($pos) == 1){$DigPos = ord($pos)%64;}
																	if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
																	if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
																	}last;}
															elsif(eof){
																if(exists($hash_pin{$DigPin[1]})){
																if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Contact_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
																if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1]."\n* Contact_Test", $format_data);
																	($pos) = $DigPin[0] =~ /^\D+/g;
																	if (length($pos) == 1){$DigPos = ord($pos)%64;}
																	if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
																	if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
																	}last;}
																else{
																if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1],$format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
																if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1], $format_data);
																	($pos) = $DigPin[0] =~ /^\D+/g;
																	if (length($pos) == 1){$DigPos = ord($pos)%64;}
																	if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
																	if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
																	}last;}
																}
															}
															close BDGfile;}
														else{
															if(exists($hash_pin{$DigPin[1]})){
															if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Contact_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
															if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1]."\n* Contact_Test", $format_data);
																($pos) = $DigPin[0] =~ /^\D+/g;
																if (length($pos) == 1){$DigPos = ord($pos)%64;}
																if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
																#print $DigPin[1]."\n";
																if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
																}}
															else{
															if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1],$format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
															if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1], $format_data);
																($pos) = $DigPin[0] =~ /^\D+/g;
																if (length($pos) == 1){$DigPos = ord($pos)%64;}
																if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
																#print $DigPin[1]."\n";
																if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
																}}
															}
														}
									if ($lineDig =~ "\;"){
									$power-> write($rowP, 4, $Total_Pin, $format_item);
									$power-> write($rowP, 5, $Power_Pin, $format_VCC);
									$power-> write($rowP, 6, $GND_Pin, $format_GND);
									$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "*Toggle_Test")', $format_data);
									$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "*Contact_Test")', $format_data);
									$power-> write($rowP, 9, $NC_Pin, $format_NC);
									$power-> write_formula($rowP, 10, "=(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-H".($rowP+1)."-I".($rowP+1)."-J".($rowP+1).")", $format_data);
									$power-> write_formula($rowP, 11, "=(H".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
									$power-> write_formula($rowP, 12, "=(I".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
									last;}
														}
													}
												}
											}
  										}
  									close Boards;}

									###### hyperlink #####################################
									$power-> write($rowP, 0, $device, $format_hylk);  			## Excel ##

									$family = "";
  									open (SourceFile, "<$testfile");
									while($lineTF = <SourceFile>)							#reading family
									{
										chomp;
										$lineTF =~ s/(^\s+)//g;								#clear head of line spacing
										#print $lineTF;
										if (substr($lineTF,0,6) eq "family")
										{$family = $lineTF . $family;}
  									}
  									close SourceFile;
  									chomp($family);
  									$power-> write($rowP, 3, $family, $format_anno);
									$rowP++;
									last;
  								}
							elsif (eof and $foundTP == 0){
							$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
							$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							$untest-> write($rowU, 2, $lineTO, $format_anno);  ## Excel ##
							$rowU++;
							goto Next_Dev;}
							}
						close TesP;
				}
			################ testable mixed device ########################################################################################
			elsif($DevTO[1] eq $device		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\%")		|($DevTO[1] =~ $device and substr($DevTO[1],length($device),1) eq "\_")
			and substr($DevTO[1],0,length($device)) eq $device
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed > -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					#$cover = 0;
					print "			Mixed_Test  ", $DevTO[1],"\n";   #, $lineTO,"\n";
			$power-> write($rowP, 4, "-", $format_data);
			$power-> write($rowP, 5, "-", $format_data);
			$power-> write($rowP, 6, "-", $format_data);
			$power-> write($rowP, 7, "-", $format_data);
			$power-> write($rowP, 8, "-", $format_data);
			$power-> write($rowP, 9, "-", $format_data);
			$power-> write($rowP, 10, "-", $format_data);
			$power-> write($rowP, 11, "-", $format_data);
			$power-> write($rowP, 12, "-", $format_data);

				 		$testname = $DevTO[1];
						open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/(^\s+|\s+$)//g;												#clear head of line spacing
							@DevTP =  split('\"', $lineTP);
							$DevTP[1] =~ s/(^\s+|\s+$)//g;

							if($DevTP[1] eq $testname	|substr($DevTP[1],6) eq $testname	#matching test name
								and substr($DevTP[1],0,length($device)) eq $device	|substr($DevTP[1],6,length($device)) eq $device)
							# and substr($lineTP,0,4) eq "test")							#matching not skipped test name
								{
									$foundTP = 1;
								#print $lineTP,"\n"; print substr($lineTP,0,1)."\n";
								if (substr($DevTP[1],6) eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = $DevTP[1];}
								if ($DevTP[1] eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = "mixed/".$testname;}
								if (index($lineTP,"on boards")> -1)
									{$testfile = "mixed/1%".$testname;}

								if ($testfile eq $testfile_last){last;}						#ignore duplicated test name
  									$testfile_last = $testfile;
									$commentTP = "";
								if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"));}   #TP comments

									$power-> write($rowP, 0, $device, $format_data);  ## Excel ##
									$power-> write($rowP, 1, $lineTO, $format_anno);  ## Excel ##
									if (length($lineTO)	> $length_TO){$length_TO = length($lineTO); $power-> set_column(1, 1, $length_TO);}

									if(substr($lineTP,0,4) eq "test"){
									$cover = 1;
									$coverage-> write($rowC, 3, 'V', $format_togg);		#Coverage
									$power-> write($rowP, 2, $lineTP, $format_anno);  ## Excel ##
									if (length($lineTP)	> $length_TP){$length_TP = length($lineTP); $power-> set_column(2, 2, $length_TP);}
									}
									if(substr($lineTP,0,1) eq "\!"){
									$PowerUT = $PowerUT + 1;
									if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}			#Coverage
									$power-> write($rowP, 2, $lineTP, $format_anno1);  ## Excel ##
									if (length($lineTP)	> $length_TP){$length_TP = length($lineTP); $power-> set_column(2, 2, $length_TP);}
									}
									$rowP++;
									last;
  								}
							elsif (eof and $foundTP == 0){
							$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
							$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							$untest-> write($rowU, 2, $lineTO, $format_anno);  ## Excel ##
							$rowU++;
							goto Next_Dev;}
							}
						close TesP;
				}
			################ testable Bscan device ########################################################################################
			elsif(($DevTO[1] =~ $device and substr($DevTO[1],length($device),8) eq "\_connect")
				and substr($DevTO[1],0,length($device)) eq $device
          and $powered == -1
          and $scan > -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					#$cover = 0;
					$length_SNail = 10;
					print "			Bscan_Test  ", $DevTO[1],"\n";   #, $lineTO,"\n";
			$power-> write($rowP, 4, "-", $format_data);
			$power-> write($rowP, 5, "-", $format_data);
			$power-> write($rowP, 6, "-", $format_data);
			$power-> write($rowP, 7, "-", $format_data);
			$power-> write($rowP, 8, "-", $format_data);
			$power-> write($rowP, 9, "-", $format_data);
			$power-> write($rowP, 10, "-", $format_data);
			$power-> write($rowP, 11, "-", $format_data);
			$power-> write($rowP, 12, "-", $format_data);

				 		$testname = $DevTO[1];
						open (TesP, "<testplan");											#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/(^\s+|\s+$)//g;									#clear head of line spacing
							@DevTP =  split('\"', $lineTP);
							$DevTP[1] =~ s/(^\s+|\s+$)//g;

							if($DevTP[1] eq $testname	|substr($DevTP[1],8) eq $testname	#matching test name
								and substr($DevTP[1],0,length($device)) eq $device	|substr($DevTP[1],8,length($device)) eq $device)
							# and substr($lineTP,0,4) eq "test")							#matching not skipped test name
								{
									$foundTP = 1;
								#print $lineTP,"\n"; print substr($lineTP,0,1)."\n";
								if (substr($DevTP[1],8) eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = $DevTP[1];}
								if ($DevTP[1] eq $testname and index($lineTP,"on boards")== -1)
									{$testfile = "digital/".$testname;}
								if (index($lineTP,"on boards")> -1)
									{$testfile = "digital/1%".$testname;}

								if ($testfile eq $testfile_last){last;}						#ignore duplicated test name
  									$testfile_last = $testfile;
									$commentTP = "";
								if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"));}   #TP comments

									#print $lineTP,"\n";
									#print $device,"\n";
									$power-> write($rowP, 1, $lineTO, $format_anno);  				## Excel ##
									if (length($lineTO)	> $length_TO){$length_TO = length($lineTO); $power-> set_column(1, 1, $length_TO);}

									if(substr($lineTP,0,4) eq "test"){
									$cover = 1;
									$coverage-> write($rowC, 4, 'V', $format_togg);					#Coverage
									$power-> write($rowP, 2, $lineTP, $format_anno);  				## Excel ##
									if (length($lineTP)	> $length_TP){$length_TP = length($lineTP); $power-> set_column(2, 2, $length_TP);}
									}
									if(substr($lineTP,0,1) eq "\!"){
									$PowerUT = $PowerUT + 1;
									if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_NC);}		#Coverage
									$power-> write($rowP, 2, $lineTP, $format_anno1);  				## Excel ##
									if (length($lineTP)	> $length_TP){$length_TP = length($lineTP); $power-> set_column(2, 2, $length_TP);}
									}
									###### hyperlink #####################################
									$power-> write_url($rowP, 0, 'internal:'.$device.'!A1');    	## hyperlink

									if ($worksheet == 0){
									$worksheet = 1;
									$IC = $bom_coverage_report-> add_worksheet($device);			## hyperlink
									$IC-> write_url('A1', 'internal:PowerTest!A1');  				## hyperlink
									$IC->conditional_formatting('A1:GR999',
									    {
									    	type     => 'text',
									     	criteria => 'containing',
									     	value    => 'Toggle_Test',
									     	format   => $format_togg,
									    });

									$IC->conditional_formatting('A1:GR999',
									    {
									    	type     => 'text',
									     	criteria => 'containing',
									     	value    => 'Contact_Test',
									     	format   => $format_pin,
									    });

									if($length_SNail < 10){$length_SNail = 10;}
									open (Boards, "< board");
									while($lineDig = <Boards>)
									{
										$lineDig =~ s/(^\s+|\s+$)//g;
										#print $lineDig;
										if (substr($lineDig,0,7) eq "DEVICES")
											{
											while($lineDig = <Boards>)								
												{
												#chomp($lineDig);
												$lineDig =~ s/(^\s+|\s+$)//g;
												$device1 = uc($device);
												#print $lineDig."\n";
												if ($lineDig eq $device1)
													{while($lineDig = <Boards>) 
														{
														$Total_Pin++;
														@BscanNail = split('\.',$lineDig);
														$BscanNail[0] =~ s/(^\s+|\s+$)//g;
														$BscanNail[1] =~ s/(^\s+|\s+$)//g;
														if ($BscanNail[1] =~ /(GND|GROUND)/){
															if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_GND); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
															if ($BscanNail[0] =~ /^\D/i){$IC-> write($BscanNail[0], $BscanNail[1], $format_GND);
															($pos) = $BscanNail[0] =~ /^\D+/g;
															if (length($pos) == 1){$NailPos = ord($pos)%64;}
															if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
															if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
															}
															$GND_Pin++;
														}
														elsif ($BscanNail[1] =~ /(^\+0|^0V|^\+1|^1V|^\+2|^2V|^\+3|^3V|^\+5|^5V|^V_|^VCC|^VDD|^PP|^P0V|^P1V|^P2V|^P3V|^P5V)/){
															if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_VCC); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
															if ($BscanNail[0] =~ /^\D/i){$IC-> write($BscanNail[0], $BscanNail[1], $format_VCC);
															($pos) = $BscanNail[0] =~ /^\D+/g;
															if (length($pos) == 1){$NailPos = ord($pos)%64;}
															if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
															if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
															}
															$Power_Pin++;
														}
														elsif ($BscanNail[1] =~ /(^NC_|_NC$|NONE)/){
															if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_NC); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
															if ($BscanNail[0] =~ /^\D/i){$IC-> write($BscanNail[0], $BscanNail[1], $format_NC);
															($pos) = $BscanNail[0] =~ /^\D+/g;
															if (length($pos) == 1){$NailPos = ord($pos)%64;}
															if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
															if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
															}
															$NC_Pin++;
														}
														else{
															if(exists($hash_pin{$BscanNail[1]})	){
															if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1]."\n* Contact_Test", $format_data); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
															if ($BscanNail[0] =~ /^\D/i){$IC-> write($BscanNail[0], $BscanNail[1]."\n* Contact_Test", $format_data);
															($pos) = $BscanNail[0] =~ /^\D+/g;
															if (length($pos) == 1){$NailPos = ord($pos)%64;}
															if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
															if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
															}}
															else{
															if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_data); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
															if ($BscanNail[0] =~ /^\D/i){$IC-> write($BscanNail[0], $BscanNail[1], $format_data);
															($pos) = $BscanNail[0] =~ /^\D+/g;
															if (length($pos) == 1){$NailPos = ord($pos)%64;}
															if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
															if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
															}}
														}
									if ($lineDig =~ "\;"){
									$power-> write($rowP, 4, $Total_Pin, $format_item);
									$power-> write($rowP, 5, $Power_Pin, $format_VCC);
									$power-> write($rowP, 6, $GND_Pin, $format_GND);
									$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "*Toggle_Test")', $format_data);
									$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "*Contact_Test")', $format_data);
									$power-> write($rowP, 9, $NC_Pin, $format_NC);
									$power-> write_formula($rowP, 10, "=(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-H".($rowP+1)."-I".($rowP+1)."-J".($rowP+1).")", $format_data);
									$power-> write_formula($rowP, 11, "=(H".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
									$power-> write_formula($rowP, 12, "=(I".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
									last;}
														}
													}
												}
											}
  										}
  									close Boards;}

									###### hyperlink #####################################
									$power-> write($rowP, 0, $device, $format_hylk);  				## Excel ##
									
									$family = "";
  									open (SourceFile, "<$testfile");
									while($lineTF = <SourceFile>)							#reading family
									{
										chomp;
										$lineTF =~ s/^ +//;                               	#clear head of line spacing
										#print $lineTF;
										if (substr($lineTF,0,6) eq "family")
										{$family = $lineTF . $family;}
  										
  										if (substr($lineTF,0,5) eq "nodes")
											{while($lineTF = <SourceFile>)
												{
													$lineTF =~ s/(^\s+|\s+$)//g;
													if ($lineTF =~ "end nodes"){last;}
													if (substr($lineTF,0,1) ne "\!") 
													{
														#print $lineTF."\n";
														@BscanNail = split('\"',$lineTF);
														#print $BscanNail[1]."\n";
														#print $BscanNail[3]."\n";
														if ($BscanNail[1] =~ "\%"){$BscanNail[1] = substr($BscanNail[1],2);}
														$BscanPin = substr($BscanNail[3], index($BscanNail[3],"\.")+1);
														#print $BscanPin."\n";
														if ($BscanPin =~ /^\d/){$IC-> write(int($BscanPin)-1, 0, $BscanNail[1]."\n* Toggle_Test", $format_data);}
														if ($BscanPin =~ /^\D/i){$IC-> write($BscanPin, $BscanNail[1]."\n* Toggle_Test", $format_data);}
														$Toggle_Pin++;
														}
													}
												}
  										}
  									close SourceFile;
  									chomp($family);
  									$power-> write($rowP, 3, $family, $format_anno);
									$rowP++;
									last;
  								}
							elsif (eof and $foundTP == 0){
							$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
							$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							$untest-> write($rowU, 2, $lineTO, $format_anno);  ## Excel ##
							$rowU++;
							goto Next_Dev;}
							}
						close TesP;
				}
      ################ reservation ########################################################################################################
      elsif (eof and $foundTO == 0)
      	{
      		print "			NO Test Item Found\n"; 
			$coverage-> write($rowC, 5, 'N', $format_NC);			#Coverage
			$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
			$untest-> write($rowU, 1, "NO test item found in TestOrder.", $format_data);  ## Excel ##
			$untest-> write($rowU, 2, "Check TJ/SP testing.", $format_anno);  ## Excel ##
			$rowU++;
      		goto Next_Dev;
      		}
      #####################################################################################################################################
			}
Next_Dev:
}
$row = 5; $col = 1;	$summary-> write($row, $col, $PowerUT, $format_data);

############################### shorts threshold statistic ################################################################################

print  "\n  >>> Analyzing shorts threshold ...\n";

$node = 1;

open (Thres, "< shorts") || open (Thres, "< 1%shorts"); 
	while($nodes = <Thres>)
	{
		chomp $nodes;
		$nodes =~ s/^ +//;	   #clear head of line spacing
		if ($nodes =~ "threshold") 
			{
				$thres = substr($nodes, index($nodes,"threshold")+10);
				if ($nodes =~ "\!"){$thres = substr($nodes, 10, index($nodes,"\!")-10);}
				$thres =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			}
		if ($nodes =~ "nodes")
		{
			if(substr($nodes,0,1) eq "!"){
			$short_thres-> write($node, 0, substr($nodes, 0, rindex($nodes,"!")), $format_data);  ## Nodes ##
			$short_thres-> write($node, 1, substr($nodes, rindex($nodes,"!")), $format_data);  ## Thres ##
				}
			elsif(substr($nodes,0,5) eq "nodes"){
			$short_thres-> write($node, 0, $nodes, $format_data);  ## Nodes ##
			$short_thres-> write($node, 1, $thres, $format_data);  ## Thres ##
				}
			$node++;
			#print $nodes."\n";
			}
		}
close Thres;

########################################################@##################################################################################


$bom_coverage_report->close();

print  "\n  >>> Completed ...\n";

END_Prog:

$end_time = time();
$duration = $end_time - $start_time;
printf "\n  runtime: %.4f Sec\n", $duration;

print "\n";
system 'pause';
exit;

