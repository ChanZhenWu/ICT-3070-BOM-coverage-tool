#!/usr/bin/perl
print "\n";
print "*******************************************************************************\n";
print "  Bom Coverage ckecking tool for 3070 <v5.99>\n";
print "  Author: Noon Chen\n";
print "  A Professional Tool for Test.\n";
print "  ",scalar localtime;
print "\n*******************************************************************************\n";
print "\n";

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
$row = 5; $col = 1;	$summary-> write($row, $col, '=COUNTIF(PowerTest!C2:C99999,"Skipped - *")', $format_data);
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

print "  please specify BOM list file: ";
   $bom=<STDIN>;
   chomp $bom;

$start_time = time();

##################### loading bom to hash ################################################
print "  gethering all BOM devices...";
my @bom_list = ();
open (Bom, "< $bom"); 
	while($dev = <Bom>)
	{
		$dev =~ s/(^\s+|\s+$)//g;
		$dev = lc($dev);

		$bom_list = push(@bom_list, $dev);
		#print $bom_list,"	--	",$value."\n";
	}
close Bom;
print "[DONE]\n";

$length = scalar @bom_list;
print "	bom: ", $length,"\n";

##################### loading BDG to hash ################################################
my %bdg_list = ();

my $BDG_File = "./bdg_data/dig_inc_ver_fau.dat";
	if(-e $BDG_File){
	open (BDGFile, "< ./bdg_data/dig_inc_ver_fau.dat");
	while($bdg = <BDGFile>)
	{
		$bdg =~ s/(^\s+|\s+$)//g;
		@bdg = split('\"',$bdg);

		if($bdg[0] =~ "Verifying faults on"){
			if ($bdg[1] =~ "\%"){$bdg[1] = substr($bdg[1],2);}
			if ($bdg[3] =~ "\%"){$bdg[3] = substr($bdg[3],2);}
			$bdg[1] = uc($bdg[1]);
			#print $bdg[1],"\n";
			$bdg[3] = uc($bdg[3]);
			#print $bdg[3],"\n";
			$bdg_list{$bdg[1]} = $bdg[3];
			#print $bdg[1]," -- ",$bdg[3]."\n";
			}
		}
	close BDGFile;
	}

@keysBDG = keys %bdg_list;
$sizeBDG = @keysBDG;
print "	BDG: ", $sizeBDG,"\n";

##################### loading pins to hash ###############################################
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

@keys = keys %hash_pin;
$size = @keys;
print "	pins: ".$size."\n";

##################### loading testorder to hash ##########################################
my %testorder = ();
open (TO, "< testorder"); 
	while($dev = <TO>)
	{
		$value = '';
		@dev = split('\"',$dev);
		$dev[0] =~ s/(^\s+|\s+$)//g;
		$dev[1] =~ s/(^\s+|\s+$)//g;
		$dev[2] = substr($dev[2],1); $dev[2] =~ s/(^\s+|\s+$)//g;
		if ($dev[0] eq "test resistor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-res'}
		if ($dev[0] eq "test capacitor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-cap'}
		if ($dev[0] eq "test jumper" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-jmp'}
		if ($dev[0] eq "test diode" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-dio'}
		if ($dev[0] eq "test zener" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-zen'}
		if ($dev[0] eq "test inductor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-ind'}
		if ($dev[0] eq "test fuse" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-fuse'}

		if ($dev[0] =~ "resistor" and $dev[2] eq "nulltest"){$value = 'untest-res'}
		if ($dev[0] =~ "capacitor" and $dev[2] eq "nulltest"){$value = 'untest-cap'}
		if ($dev[0] =~ "jumper" and $dev[2] eq "nulltest"){$value = 'untest-jmp'}
		if ($dev[0] =~ "diode" and $dev[2] eq "nulltest"){$value = 'untest-dio'}
		if ($dev[0] =~ "zener" and $dev[2] eq "nulltest"){$value = 'untest-zen'}
		if ($dev[0] =~ "inductor" and $dev[2] eq "nulltest"){$value = 'untest-ind'}
		if ($dev[0] =~ "fuse" and $dev[2] eq "nulltest"){$value = 'untest-fuse'}

		if ($dev[0] =~ "resistor" and $dev[2] =~ "nulltest" and $dev[2] =~ "\!"){$value = 'paral-res'}
		if ($dev[0] =~ "capacitor" and $dev[2] =~ "nulltest" and $dev[2] =~ "\!"){$value = 'paral-cap'}
		if ($dev[0] =~ "jumper" and $dev[2] =~ "nulltest" and $dev[2] =~ "\!"){$value = 'paral-jmp'}
		if ($dev[0] =~ "diode" and $dev[2] =~ "nulltest" and $dev[2] =~ "\!"){$value = 'paral-dio'}
		if ($dev[0] =~ "zener" and $dev[2] =~ "nulltest" and $dev[2] =~ "\!"){$value = 'paral-zen'}
		if ($dev[0] =~ "inductor" and $dev[2] =~ "nulltest" and $dev[2] =~ "\!"){$value = 'paral-ind'}
		if ($dev[0] =~ "fuse" and $dev[2] =~ "nulltest" and $dev[2] =~ "\!"){$value = 'paral-fuse'}

		if ($dev[0] eq "test mixed"){$value = 'tested-mix'}
		if ($dev[0] eq "test digital"){$value = 'tested-dig'}
		if ($dev[0] eq "test analog powered"){$value = 'tested-pwr'}
		if ($dev[0] eq "test scan connect"){$value = 'tested-bscan'}

		if ($dev[0] eq "skip mixed"){$value = 'untest-mix'}
		if ($dev[0] eq "skip digital"){$value = 'untest-dig'}
		if ($dev[0] eq "skip analog powered"){$value = 'untest-pwr'}
		if ($dev[0] eq "skip scan connect"){$value = 'untest-bscan'}

		$testorder{$dev[1]} = $value;
		# print $dev[1],"	--	",$value."\n";
	}
close TO;

@keysTO = keys %testorder;
$sizeTO = @keysTO;
print "	testorder: ".$sizeTO."\n";

##################### loading testplan to hash ###########################################
%testplan = ();
$learn = 0;
open (TP, "< testplan"); 
	while($dev = <TP>)
	{
		$value = '';
		@dev = split('\"',$dev);
		$dev[0] =~ s/(^\s+|\s+$)//g;
		$dev[1] =~ s/(^\s+|\s+$)//g;
		$dev[2] =~ s/(^\s+|\s+$)//g;
		if ($dev =~ "learn capacitance on"){$learn = 1;}
		if ($dev =~ "learn capacitance off"){$learn = 0;}
		next if ($learn == 1);
		
		if ($dev[0] eq "test"){
			if($dev[1] =~ 'mixed/'){$dev[1] = substr($dev[1],6)}
			if($dev[1] =~ 'analog/'){$dev[1] = substr($dev[1],7)}
			if($dev[1] =~ 'digital/'){$dev[1] = substr($dev[1],8)}
			if($dev[2] =~ "\!"){$value = "tested".substr($dev[2],rindex($dev[2],"\!") + 1);}
			if($dev[2] !~ "\!"){$value = "tested";}
		}
		elsif ($dev[0] =~ "test" and substr($dev[0],0,1) eq "\!"){
			if($dev[1] =~ 'mixed/'){$dev[1] = substr($dev[1],6)}
			if($dev[1] =~ 'analog/'){$dev[1] = substr($dev[1],7)}
			if($dev[1] =~ 'digital/'){$dev[1] = substr($dev[1],8)}
			if($dev[2] =~ "\!"){$value = "skipped-".substr($dev[2],rindex($dev[2],"\!") + 1);}
			if($dev[2] !~ "\!"){$value = "skipped-";}
		}
		$testplan{$dev[1]} = $value;
		#print $dev[1],"	--	",$value."\n";
		#print substr($testplan{$dev[1]},0,6),"\n";
	}
close TP;

@keysTP = keys %testplan;
$sizeTP = @keysTP;
print "	testplan: ".$sizeTP."\n";

##########################################################################################
print "\n";
foreach $device (@bom_list)
{
	$Total_Pin = 0;
	$Power_Pin = 0;
	$GND_Pin = 0;
	$Toggle_Pin = 0;
	$NC_Pin = 0;
	#-------------------------------------------------------------------------------------
	$worksheet = 0;
	$foundTO = 0;
	$foundTP = 0;
	$device =~ s/(^\s+|\s+$)//g;                     #clear all spacing
	$device = lc($device);

	$rowC = $rowC+1;	$cover = 0;	$UNCover = 0;
	$coverage-> write($rowC, 0, $device, $format_data);			#Coverage
	@array = ("-","-","-","-","-");
		$array_ref = \@array;
		$coverage-> write_row($rowC, 1, $array_ref, $format_data);
	$len_device = length($device);
	print "	Analyzing ", $device, " .....\n";

	################ testable device ##############################################################################################
	if(substr($testplan{$device},0,6) eq "tested"
		and ($testorder{$device} eq "tested-res"
		or $testorder{$device} eq "tested-cap"
		or $testorder{$device} eq "tested-jmp"
		or $testorder{$device} eq "tested-dio"
		or $testorder{$device} eq "tested-zen"
		or $testorder{$device} eq "tested-ind"
		or $testorder{$device} eq "tested-fuse")
       )
		{
			#$foundTO = 1;
			print "			General AnaTest		", $device."\n";   #, $lineTO,"\n";
			@array = ("-","-","-","-");
				$array_ref = \@array;
				$tested-> write_row($rowT, 2, $array_ref, $format_data);
		
			$UNCover = 1;
			$coverage-> write($rowC, 1, 'V', $format_togg);			#Coverage

				open(SourceFile, "<analog/$device")||open(SourceFile, "<analog/1%$device");
					while($lineTF = <SourceFile>)							#read parameter
					{
						$len = 0;
						chomp;
						$lineTF =~ s/^ +//;                               	#clear head of line spacing
						#print $lineTF;
						if (substr($lineTF,0,8) eq "resistor")				#### matching resistor ########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,8), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,9));
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
							if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
							if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
							if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$device},6), $format_anno);  	## Comment ##
							$rowT++;
							goto Next_Dev;
						}
						elsif (substr($lineTF,0,9) eq "capacitor")				#### matching capacitor ########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,9), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,10));
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
							if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
							if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
							if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$device},6), $format_anno);  	## Comment ##
							$rowT++;
							goto Next_Dev;
						}
						elsif (substr($lineTF,0,8) eq "inductor")					#### matching inductor ########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,8), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,9));
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
							if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
							if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
							if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$device},6), $format_anno);  	## Comment ##
							$rowT++;
							goto Next_Dev;
						}
						elsif (substr($lineTF,0,5) eq "diode")					#### matching diode ######
						{
							$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,5), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,6));
							if ($lineTF !~ m/\"/g){
							#$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							$tested-> write($rowT, 3, $param[0], $format_data);  						## HiLimit ##
							$tested-> write($rowT, 4, $param[1], $format_data);	  						## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$device},6), $format_anno);  	## Comment ##
							}
							if ($lineTF =~ m/\"/g){
							$DioNom = "";	$DioNom =  $DioNom . $param[0];  							## Nominal ## 
							$DioHiL = "";	$DioHiL =  $DioHiL . $param[1];  							## HiLimit ##
							$DioLoL = "";	$DioLoL =  $DioLoL . $param[2];	  							## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$device},6), $format_anno);  	## Comment ##
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
							#goto Next_Dev;
						}
						elsif (substr($lineTF,0,5) eq "zener")					 ####matching zener##
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,5), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,6));
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
							if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
							if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
							if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$device},6), $format_anno);  	## Comment ##
							$rowT++;
							goto Next_Dev;
						}
						elsif (substr($lineTF,0,6) eq "jumper")						####matching jumper########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,6), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,6));
							$OP = 0; if ($lineTF =~ "op") {$OP = 1;}
							if($OP == 0){$tested-> write($rowT, 3, $param[0], $format_data);	$tested-> write($rowT, 5, substr($testplan{$device},6), $format_anno);}  ## Excel ##
							if($OP == 1){$tested-> write($rowT, 4, $param[0], $format_STP);	$tested-> write($rowT, 5, "OP test", $format_STP);}	## Excel ##
							$rowT++;
							goto Next_Dev;
						}
						elsif (substr($lineTF,0,4) eq "fuse")					  	####matching fuse########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,6), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,6));
							$tested-> write($rowT, 3, $param[0], $format_data);
							$tested-> write($rowT, 5, substr($testplan{$device},6), $format_anno);  ## Excel ##
							$rowT++;
							goto Next_Dev;
						}
						elsif (eof and $foundTO == 1)					  			####no parameter######
						{
						$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
						$untest-> write($rowU, 1, "No Test Parameter Found in TestFile.", $format_data);  	## Excel ##
						$rowU++;
						goto Next_Dev;
						}}}
	################ testplan skipped devices ########################################################################################
	elsif (substr($testplan{$device},0,7) eq "skipped"
			and ($testorder{$device} eq "tested-res"
			or $testorder{$device} eq "tested-cap"
			or $testorder{$device} eq "tested-jmp"
			or $testorder{$device} eq "tested-dio"
			or $testorder{$device} eq "tested-zen"
			or $testorder{$device} eq "tested-ind"
			or $testorder{$device} eq "tested-fuse")
    		)
		{
			$foundTO = 1;
			print "			General SkipTest	", $device,"\n";   #, $lineTO,"\n";
			if ($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_NC);}			#Coverage
			$untest-> write($rowU, 0, $device, $format_data);		## Excel ##
			$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_anno1);  ## Excel ##
			$untest-> write($rowU, 2, substr($testplan{$device},8), $format_anno);  ## Excel ##
			$rowU++;
			next; #goto Next_Dev;
		}
	################ testplan inexistent devices #####################################################################################
	elsif (substr($testplan{$device},0,6) eq ""
			and ($testorder{$device} eq "tested-res"
			or $testorder{$device} eq "tested-cap"
			or $testorder{$device} eq "tested-jmp"
			or $testorder{$device} eq "tested-dio"
			or $testorder{$device} eq "tested-zen"
			or $testorder{$device} eq "tested-ind"
			or $testorder{$device} eq "tested-fuse")
    		)
		{
			$foundTO = 1;
			print "			Testplan InExistent	", $device,"\n";   #, $lineTO,"\n";
			if ($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_NC);}				#Coverage
			$untest-> write($rowU, 0, $device, $format_data);								## Excel ##
			$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);	## Excel ##
			$untest-> write($rowU, 2, $testorder{$device}, $format_anno);					## Excel ##
			$rowU++;
			next; #goto Next_Dev;
		}
	################ testorder skipped devices #######################################################################################
	elsif($testorder{$device} eq "untest-res"
		or $testorder{$device} eq "untest-cap"
		or $testorder{$device} eq "untest-jmp"
		or $testorder{$device} eq "untest-dio"
		or $testorder{$device} eq "untest-zen"
		or $testorder{$device} eq "untest-ind"
		or $testorder{$device} eq "untest-fuse"
       )
		{
			if ($testorder{$device} ne "untest-dio"){$foundTO = 1;}
			print "			General NullTest	", $device,"\n";   #, $lineTO,"\n";
			$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
			$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_anno);  ## Excel ##
		$UTline = "";
		$fileF = 0;
		
		if($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_NC);}			#Coverage
		open(ALL, "<analog/$device") || open(ALL, "<analog/1%$device") || $untest-> write($rowU, 2, "!TestFile not found.", $format_anno1);
		while($line = <ALL>)
			{
			$fileF = 1;
			if (index($line,$device)>1){
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;}
			elsif (eof){last;}
			}
			$UTline =~ s/(^\s+|\s+$)//g;
			if($UTline eq "" and $fileF == 1){$untest-> write($rowU, 2, "No Comments Found in TestFile.", $format_anno1);}
			if($UTline ne ""){
			$untest-> write($rowU, 2, $UTline, $format_anno);
			$untest-> set_column(2, 2, $length_anno);
			}
			$rowU++;
			close ALL;
			#next; #goto Next_Dev;
		}
	################ parallel tested devices ######################################################################################
	elsif($testorder{$device} eq "paral-res"
		or $testorder{$device} eq "paral-cap"
		or $testorder{$device} eq "paral-jmp"
		or $testorder{$device} eq "paral-dio"
		or $testorder{$device} eq "paral-zen"
		or $testorder{$device} eq "paral-ind"
		or $testorder{$device} eq "paral-fuse"
       )
		{
			$foundTO = 1;
			print "			General ParalTest	", $device,"\n";		#, $lineTO,"\n";
			
			if($UNCover == 0){$coverage-> write($rowC, 1, 'L', $format_pin);}			#Coverage
			$limited-> write($rowL, 0, $device, $format_data);		## Excel ##
			#$limited-> write($rowL, 1, $dev[2], $format_anno);		## Excel ##
		$UTline = "";
		$fileF = 0;
		
		open(ALL, "<analog/$device") || open(ALL, "<analog/1%$device") || $limited-> write($rowL, 1, "!TestFile not found.", $format_anno1);
		while($line = <ALL>)
			{
			$fileF = 1;
			if (index($line,$device)>1){
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;}
			elsif (eof){last;}
			}
			$UTline =~ s/(^\s+|\s+$)//g;
			if($UTline eq "" and $fileF == 1){$limited-> write($rowL, 1, "No Comments Found in TestFile.", $format_anno1);}
			if($UTline ne ""){
			$limited-> write($rowL, 1, $UTline, $format_anno);	
			$limited-> set_column(1, 1, $length_anno);
			}
			$rowL++;
			close ALL;
			goto Next_Dev;
		}
	################ testable analog powered test #################################################################################
	elsif($testorder{$device} eq "tested-pwr" or $testorder{$device} eq "untest-pwr"
       )
		{
			$foundTO = 0;
			print "			General PwrTest		", $device,"\n";   #, $lineTO,"\n";

			if($testorder{$device} eq "tested-pwr" and substr($testplan{$device},0,6) eq "tested"){
				$cover = 1;
				$coverage-> write($rowC, 3, 'V', $format_togg);								#Coverage

				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

				$power-> write($rowP, 0, $device, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$device, $format_anno);				## Excel ##
				$rowP++;
				}
			elsif($testorder{$device} eq "tested-pwr" and substr($testplan{$device},0,7) eq "skipped"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}				#Coverage

				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

				$power-> write($rowP, 0, $device, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);		## Excel ##
				$rowP++;
				}
			elsif($testorder{$device} eq "untest-pwr"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}				#Coverage
				$untest-> write($rowU, 0, $device, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  				## Excel ##
				$rowU++;
				}
			elsif($testorder{$device} eq "tested-pwr"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}				#Coverage
				$untest-> write($rowU, 0, $device, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  				## Excel ##
				$rowU++;
				}
			#goto Next_Dev;
		}
	################ testable digital test ########################################################################################
	elsif($testorder{$device} eq "tested-dig" or $testorder{$device} eq "untest-dig"
       )
		{
			$foundTO = 0;
			$length_DigPin = 10;
			print "			General DigiTest	", $device,"\n";   #, $lineTO,"\n";
			
			if($testorder{$device} eq "tested-dig" and substr($testplan{$device},0,6) eq "tested"){
				$cover = 1;
				$coverage-> write($rowC, 2, 'V', $format_togg);								#Coverage
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$device, $format_anno);				## Excel ##
				}
			elsif($testorder{$device} eq "tested-dig" and substr($testplan{$device},0,7) eq "skipped"){
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_NC);}				#Coverage
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);			## Excel ##
				}
			if($testorder{$device} eq "tested-dig" and $testplan{$device},){
			
				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

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
										if (exists($bdg_list{uc($device)."\.".$DigPin[0]})){
										#print uc($device)."\.".$DigPin[0],"\n";
										if($bdg_list{uc($device)."\.".$DigPin[0]})
											{
											#print uc($device)."\.".$DigPin[0],"\n";
											if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Toggle_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
											if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1]."\n* Toggle_Test", $format_data);
												($pos) = $DigPin[0] =~ /^\D+/g;
												if (length($pos) == 1){$DigPos = ord($pos)%64;}
												if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
												if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
												}}
											}
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
												if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
												}}
											}
										}
									if ($lineDig =~ "\;"){
									$power-> write($rowP, 4, $Total_Pin, $format_item);
									$power-> write($rowP, 5, $Power_Pin, $format_VCC);
									$power-> write($rowP, 6, $GND_Pin, $format_GND);
									$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "* Toggle_Test*")', $format_data);
									$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "* Contact_Test*")', $format_data);
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
  				open(SourceFile, "<digital/$device")||open(SourceFile, "<digital/1%$device");
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
				}
			elsif($testorder{$device} eq "untest-dig"){
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_NC);}				#Coverage
				$untest-> write($rowU, 0, $device, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  				## Excel ##
				$rowU++;
				}
			elsif($testorder{$device} eq "tested-dig"){
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_NC);}				#Coverage
				$untest-> write($rowU, 0, $device, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  				## Excel ##
				$rowU++;
				}
		}
	################ testable mixed test ##########################################################################################
	elsif($testorder{$device} eq "tested-mix" or $testorder{$device} eq "untest-mix"
       )
		{
			$foundTO = 0;
			print "			General MixTest		", $device,"\n";   #, $lineTO,"\n";
			
			if($testorder{$device} eq "tested-mix" and substr($testplan{$device},0,6) eq "tested"){
				$cover = 1;
				$coverage-> write($rowC, 3, 'V', $format_togg);								#Coverage

				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

				$power-> write($rowP, 0, $device, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$device, $format_anno);				## Excel ##
				$rowP++;
				}
			elsif($testorder{$device} eq "tested-mix" and substr($testplan{$device},0,7) eq "skipped"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}				#Coverage

				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

				$power-> write($rowP, 0, $device, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);		## Excel ##
				$rowP++;
				}
			elsif($testorder{$device} eq "untest-mix"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}				#Coverage
				$untest-> write($rowU, 0, $device, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  				## Excel ##
				$rowU++;
				}
			elsif($testorder{$device} eq "tested-mix"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}				#Coverage
				$untest-> write($rowU, 0, $device, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  				## Excel ##
				$rowU++;
				}
			#goto Next_Dev;
		}
	################ testable Bscan device ########################################################################################
	elsif($testorder{$device} eq "tested-bscan" or $testorder{$device} eq "untest-bscan"
       )
		{
			$foundTO = 0;
			$length_SNail = 10;
			print "			General BscTest	", $device,"\n";   #, $lineTO,"\n";
			
			if($testorder{$device} eq "tested-bscan" and substr($testplan{$device},0,6) eq "tested"){
				$cover = 1;
				$coverage-> write($rowC, 4, 'V', $format_togg);								#Coverage
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$device, $format_anno);				## Excel ##
				}
			elsif($testorder{$device} eq "tested-bscan" and substr($testplan{$device},0,7) eq "skipped"){
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_NC);}				#Coverage
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);			## Excel ##
				}
			if($testorder{$device} eq "tested-bscan" and $testplan{$device},){
			
				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

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
									$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "* Toggle_Test*")', $format_data);
									$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "* Contact_Test*")', $format_data);
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
  				open(SourceFile, "<digital/$device")||open(SourceFile, "<digital/1%$device");
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
  				}
			elsif($testorder{$device} eq "untest-bscan"){
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_NC);}				#Coverage
				$untest-> write($rowU, 0, $device, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  				## Excel ##
				$rowU++;
				}
			elsif($testorder{$device} eq "tested-bscan"){
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_NC);}				#Coverage
				$untest-> write($rowU, 0, $device, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  				## Excel ##
				$rowU++;
				}		
		}

   #~~~~~~~~~~~~~~~ Multiple Test ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if ($foundTO == 0
		or $testorder{$device} eq "tested-mix" or $testorder{$device} eq "untest-mix"
		or $testorder{$device} eq "tested-pwr" or $testorder{$device} eq "untest-pwr"
		or $testorder{$device} eq "tested-dig" or $testorder{$device} eq "untest-dig"
		or $testorder{$device} eq "tested-bscan" or $testorder{$device} eq "untest-bscan"
		or $testorder{$device} eq "tested-dio" or $testorder{$device} eq "untest-dio")
		{
		#print "--0--",$testorder{$device},"\n";
		#print $keysTO[11],"\n";
		my $Mult_file = '';
		for ($i = 0; $i <= $sizeTO; $i = $i + 1){
			if ($keysTO[$i] =~ $device and substr($keysTO[$i],length($device),1) eq "\%" |substr($keysTO[$i],length($device),1) eq "\_" ){
				#print "--1--",$testorder{$device},"\n";
				#print substr($keysTO[$i],length($device),1),"\n";
				#print $keysTO[$i],'	--	',$testorder{$keysTO[$i]},"\n";
				#print substr($testplan{$keysTO[$i]},0,6),"\n";
				$Mult_file = $keysTO[$i];
				#print $Mult_file,"\n";
	#%%%%%%%%%%%%%%% testable device %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
				if(substr($testplan{$keysTO[$i]},0,6) eq "tested"
					and ($testorder{$keysTO[$i]} eq "tested-res"
					or $testorder{$keysTO[$i]} eq "tested-cap"
					or $testorder{$keysTO[$i]} eq "tested-jmp"
					or $testorder{$keysTO[$i]} eq "tested-dio"
					or $testorder{$keysTO[$i]} eq "tested-zen"
					or $testorder{$keysTO[$i]} eq "tested-ind"
					or $testorder{$keysTO[$i]} eq "tested-fuse")
				   )
				   	{
				#$foundTO = 1;
				print "			Multiple AnaTest	", $Mult_file."\n";   #, $lineTO,"\n";
				@array = ("-","-","-","-");
				$array_ref = \@array;
				$tested-> write_row($rowT, 2, $array_ref, $format_data);
				
				$UNCover = 1;	
				$coverage-> write($rowC, 1, 'V', $format_togg);			#Coverage

				open(SourceFile, "<analog/$Mult_file")||open(SourceFile, "<analog/1%$Mult_file");
					while($lineTF = <SourceFile>)							#read parameter
					{
						$len = 0;
						chomp;
						$lineTF =~ s/^ +//;                               	#clear head of line spacing
						#print $lineTF;
						if (substr($lineTF,0,8) eq "resistor")				#### matching resistor ########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,8), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,9));
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
							if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
							if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
							if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
							$rowT++;
							last; #goto Next_Dev;
						}
						elsif (substr($lineTF,0,9) eq "capacitor")				#### matching capacitor ########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,9), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,10));
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
							if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
							if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
							if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
							$rowT++;
							last; #goto Next_Dev;
						}
						elsif (substr($lineTF,0,8) eq "inductor")					#### matching inductor ########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,8), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,9));
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
							if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
							if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
							if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
							$rowT++;
							last; #goto Next_Dev;
						}
						elsif (substr($lineTF,0,5) eq "diode")					#### matching diode ######
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,5), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,6));
							if ($lineTF !~ m/\"/g){
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							$tested-> write($rowT, 3, $param[0], $format_data);  						## HiLimit ##
							$tested-> write($rowT, 4, $param[1], $format_data);	  						## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
							}
							if ($lineTF =~ m/\"/g){
							$DioNom = "";	$DioNom =  $DioNom . $param[0];  							## Nominal ## 
							$DioHiL = "";	$DioHiL =  $DioHiL . $param[1];  							## HiLimit ##
							$DioLoL = "";	$DioLoL =  $DioLoL . $param[2];	  							## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
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
						$foundTO = 1;
							$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,5), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,6));
							$tested-> write($rowT, 2, $param[0], $format_data);  						## Nominal ## 
							if($param[1] < 40){$tested-> write($rowT, 3, $param[1], $format_data);}  	## HiLimit ##
							if($param[1] >= 40){$tested-> write($rowT, 3, $param[1], $format_STP);}  	## HiLimit ##
							if($param[2] < 40){$tested-> write($rowT, 4, $param[2], $format_data);}  	## LoLimit ##
							if($param[2] >= 40){$tested-> write($rowT, 4, $param[2], $format_STP);}		## LoLimit ##
							$tested-> write($rowT, 5, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
							$rowT++;
							last; #goto Next_Dev;
						}
						elsif (substr($lineTF,0,6) eq "jumper")						####matching jumper########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,6), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,6));
							$OP = 0; if ($lineTF =~ "op") {$OP = 1;}
							if($OP == 0){$tested-> write($rowT, 3, $param[0], $format_data);	$tested-> write($rowT, 5, substr($testplan{$Mult_file},6), $format_anno);}  ## Excel ##
							if($OP == 1){$tested-> write($rowT, 4, $param[0], $format_STP);	$tested-> write($rowT, 5, "OP test", $format_STP);}	## Excel ##
							$rowT++;
							last; #goto Next_Dev;
						}
						elsif (substr($lineTF,0,4) eq "fuse")					  	####matching fuse########
						{
						$foundTO = 1;
							$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
							$tested-> write($rowT, 1, substr($lineTF,0,6), $format_data);  				## TestType ##
							#print substr($lineTF,9)."\n";
							@param =  split('\,', substr($lineTF,6));
							$tested-> write($rowT, 3, $param[0], $format_data);
							$tested-> write($rowT, 5, substr($testplan{$Mult_file},6), $format_anno);  ## Excel ##
							$rowT++;
							last; #goto Next_Dev;
						}
						elsif (eof and $foundTO == 1)					  			####no parameter######
						{
						$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
						$untest-> write($rowU, 1, "No Test Parameter Found in TestFile.", $format_data);  	## Excel ##
						$rowU++;
						last; #goto Next_Dev;
					}}}
	#%%%%%%%%%%%%%%% testplan skipped devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif (substr($testplan{$keysTO[$i]},0,7) eq "skipped"
			and ($testorder{$keysTO[$i]} eq "tested-res"
			or $testorder{$keysTO[$i]} eq "tested-cap"
			or $testorder{$keysTO[$i]} eq "tested-jmp"
			or $testorder{$keysTO[$i]} eq "tested-dio"
			or $testorder{$keysTO[$i]} eq "tested-zen"
			or $testorder{$keysTO[$i]} eq "tested-ind"
			or $testorder{$keysTO[$i]} eq "tested-fuse")
    		)
		{
			$foundTO = 1;
			print "			Multiple SkipTest	", $Mult_file,"\n";   #, $lineTO,"\n";
			if ($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_NC);}			#Coverage
			$untest-> write($rowU, 0, $Mult_file, $format_data);		## Excel ##
			$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_anno1);  ## Excel ##
			$untest-> write($rowU, 2, substr($testplan{$Mult_file},8), $format_anno);  ## Excel ##
			$rowU++;
			#last; #goto Next_Dev;
		}
	#%%%%%%%%%%%%%%% testplan inexistent devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif (substr($testplan{$keysTO[$i]},0,6) eq ""
			and ($testorder{$keysTO[$i]} eq "tested-res"
			or $testorder{$keysTO[$i]} eq "tested-cap"
			or $testorder{$keysTO[$i]} eq "tested-jmp"
			or $testorder{$keysTO[$i]} eq "tested-dio"
			or $testorder{$keysTO[$i]} eq "tested-zen"
			or $testorder{$keysTO[$i]} eq "tested-ind"
			or $testorder{$keysTO[$i]} eq "tested-fuse")
    		)
		{
			$foundTO = 1;
			print "			Testplan InExistent	", $Mult_file,"\n";   #, $lineTO,"\n";
			if ($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_NC);}				#Coverage
			$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
			$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);	## Excel ##
			$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);				## Excel ##
			$rowU++;
			#last; #goto Next_Dev;
		}
	#%%%%%%%%%%%%%%% testorder skipped devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif($testorder{$keysTO[$i]} eq "untest-res"
		or $testorder{$keysTO[$i]} eq "untest-cap"
		or $testorder{$keysTO[$i]} eq "untest-jmp"
		or $testorder{$keysTO[$i]} eq "untest-dio"
		or $testorder{$keysTO[$i]} eq "untest-zen"
		or $testorder{$keysTO[$i]} eq "untest-ind"
		or $testorder{$keysTO[$i]} eq "untest-fuse"
       )
		{
			$foundTO = 1;
			print "			Multiple NullTest	", $Mult_file,"\n";   #, $lineTO,"\n";
			$untest-> write($rowU, 0, $Mult_file, $format_data);  ## Excel ##
			$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_anno);  ## Excel ##
		$UTline = "";
		$fileF = 0;
		
		if($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_NC);}			#Coverage
		open(ALL, "<analog/$Mult_file") || open(ALL, "<analog/1%$Mult_file") || $untest-> write($rowU, 2, "!TestFile not found.", $format_anno1);
		while($line = <ALL>)
			{
			$fileF = 1;
			if (index($line,$device)>1){
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;}
			elsif (eof){last;}
			}
			$UTline =~ s/(^\s+|\s+$)//g;
			if($UTline eq "" and $fileF == 1){$untest-> write($rowU, 2, "No Comments Found in TestFile.", $format_anno1);}
			if($UTline ne ""){
			$untest-> write($rowU, 2, $UTline, $format_anno);
			$untest-> set_column(2, 2, $length_anno);
			}
			$rowU++;
			close ALL;
			#last;
		}
	#%%%%%%%%%%%%%%% parallel tested devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif($testorder{$keysTO[$i]} eq "paral-res"
		or $testorder{$keysTO[$i]} eq "paral-cap"
		or $testorder{$keysTO[$i]} eq "paral-jmp"
		or $testorder{$keysTO[$i]} eq "paral-dio"
		or $testorder{$keysTO[$i]} eq "paral-zen"
		or $testorder{$keysTO[$i]} eq "paral-ind"
		or $testorder{$keysTO[$i]} eq "paral-fuse"
       )
		{
			$foundTO = 1;
			print "			Multiple ParalTest	", $Mult_file,"\n";		#, $lineTO,"\n";
			
			if($UNCover == 0){$coverage-> write($rowC, 1, 'L', $format_pin);}			#Coverage
			$limited-> write($rowL, 0, $Mult_file, $format_data);		## Excel ##
			#$limited-> write($rowL, 1, $dev[2], $format_anno);			## Excel ##
		$UTline = "";
		$fileF = 0;

		open(ALL, "<analog/$Mult_file") || open(ALL, "<analog/1%$Mult_file") || $limited-> write($rowL, 1, "!TestFile not found.", $format_anno1);
		while($line = <ALL>)
			{
			$fileF = 1;
			if (index($line,$device)>1){
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;}
			elsif (eof){last;}
			}
			$UTline =~ s/(^\s+|\s+$)//g;
			if($UTline eq "" and $fileF == 1){$limited-> write($rowL, 1, "No Comments Found in TestFile.", $format_anno1);}
			if($UTline ne ""){
			$limited-> write($rowL, 1, $UTline, $format_anno);	
			$limited-> set_column(1, 1, $length_anno);
			}
			$rowL++;
			close ALL;
			#last;
		}
	#%%%%%%%%%%%%%%% testable analog powered test %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif($testorder{$keysTO[$i]} eq "tested-pwr" or $testorder{$keysTO[$i]} eq "untest-pwr"
		or $testorder{$device} eq "tested-dio" or $testorder{$device} eq "untest-dio"
       )
		{
			$foundTO = 1;
			print "			Multiple PwrTest	", $Mult_file,"\n";   #, $lineTO,"\n";
			
			if($testorder{$Mult_file} eq "tested-pwr" and substr($testplan{$Mult_file},0,6) eq "tested"){
				$cover = 1;
				$coverage-> write($rowC, 3, 'V', $format_togg);								#Coverage

				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

				$power-> write($rowP, 0, $Mult_file, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$Mult_file}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);				## Excel ##
				$rowP++;
				}
			elsif($testorder{$Mult_file} eq "tested-pwr" and substr($testplan{$Mult_file},0,7) eq "skipped"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}				#Coverage

				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

				$power-> write($rowP, 0, $Mult_file, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$Mult_file}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);		## Excel ##
				$rowP++;
				}
			elsif($testorder{$Mult_file} eq "untest-pwr"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}					#Coverage
				$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  		## Excel ##
				$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);  				## Excel ##
				$rowU++;
				}
			elsif($testorder{$Mult_file} eq "tested-pwr"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}					#Coverage
				$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);  				## Excel ##
				$rowU++;
				}
		}
	#%%%%%%%%%%%%%%% testable digital test %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif($testorder{$keysTO[$i]} eq "tested-dig" or $testorder{$keysTO[$i]} eq "untest-dig"
       )
		{
			$foundTO = 1;
			$length_DigPin = 10;
			print "			Multiple DigiTest	", $Mult_file,"\n";   #, $lineTO,"\n";
			
			if($testorder{$Mult_file} eq "tested-dig" and substr($testplan{$Mult_file},0,6) eq "tested"){
				$cover = 1;
				$coverage-> write($rowC, 2, 'V', $format_togg);								#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);				## Excel ##
				}
			elsif($testorder{$Mult_file} eq "tested-dig" and substr($testplan{$Mult_file},0,7) eq "skipped"){
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_NC);}				#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);			## Excel ##
				}
			if($testorder{$Mult_file} eq "tested-dig" and $testplan{$Mult_file},){
			
				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

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
										if (exists($bdg_list{uc($device)."\.".$DigPin[0]})){
										#print uc($device)."\.".$DigPin[0],"\n";
										if($bdg_list{uc($device)."\.".$DigPin[0]})
											{
											#print uc($device)."\.".$DigPin[0],"\n";
											if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Toggle_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
											if ($DigPin[0] =~ /^\D/i){$IC-> write($DigPin[0], $DigPin[1]."\n* Toggle_Test", $format_data);
												($pos) = $DigPin[0] =~ /^\D+/g;
												if (length($pos) == 1){$DigPos = ord($pos)%64;}
												if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
												if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
												}}
											}
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
												if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
												}}
											}
										}
									if ($lineDig =~ "\;"){
									$power-> write($rowP, 4, $Total_Pin, $format_item);
									$power-> write($rowP, 5, $Power_Pin, $format_VCC);
									$power-> write($rowP, 6, $GND_Pin, $format_GND);
									$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "* Toggle_Test*")', $format_data);
									$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "* Contact_Test*")', $format_data);
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
  				open(SourceFile, "<digital/$Mult_file")||open(SourceFile, "<digital/1%$Mult_file");
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
				}
			elsif($testorder{$Mult_file} eq "untest-dig"){
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_NC);}					#Coverage
				$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  		## Excel ##
				$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);  				## Excel ##
				$rowU++;
				}
			elsif($testorder{$Mult_file} eq "tested-dig"){
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_NC);}					#Coverage
				$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);  				## Excel ##
				$rowU++;
				}
		}
	#%%%%%%%%%%%%%%% testable mixed device %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif($testorder{$keysTO[$i]} eq "tested-mix" or $testorder{$keysTO[$i]} eq "untest-mix"
       )
		{
			$foundTO = 1;
			print "			Multiple MixTest	", $Mult_file,"\n";   #, $lineTO,"\n";
			
			if($testorder{$Mult_file} eq "tested-mix" and substr($testplan{$Mult_file},0,6) eq "tested"){
				$cover = 1;
				$coverage-> write($rowC, 3, 'V', $format_togg);								#Coverage

				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

				$power-> write($rowP, 0, $Mult_file, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$Mult_file}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);				## Excel ##
				$rowP++;
				}
			elsif($testorder{$Mult_file} eq "tested-mix" and substr($testplan{$Mult_file},0,7) eq "skipped"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}				#Coverage

				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

				$power-> write($rowP, 0, $Mult_file, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$Mult_file}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);		## Excel ##
				$rowP++;
				}
			elsif($testorder{$Mult_file} eq "untest-mix"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}					#Coverage
				$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  		## Excel ##
				$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);  				## Excel ##
				$rowU++;
				}
			elsif($testorder{$Mult_file} eq "tested-mix"){
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_NC);}					#Coverage
				$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);  				## Excel ##
				$rowU++;
				}
		}
	#%%%%%%%%%%%%%%% testable Bscan device %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif($testorder{$keysTO[$i]} eq "tested-bscan" or $testorder{$keysTO[$i]} eq "untest-bscan"
       )
		{
			$foundTO = 1;
			$length_SNail = 10;
			print "			Multiple BscTest	", $Mult_file,"\n";   #, $lineTO,"\n";
			
			if($testorder{$Mult_file} eq "tested-bscan" and substr($testplan{$Mult_file},0,6) eq "tested"){
				$cover = 1;
				$coverage-> write($rowC, 4, 'V', $format_togg);								#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);				## Excel ##
				}
			elsif($testorder{$Mult_file} eq "tested-bscan" and substr($testplan{$Mult_file},0,7) eq "skipped"){
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_NC);}				#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);			## Excel ##
				}
			if($testorder{$Mult_file} eq "tested-bscan" and $testplan{$Mult_file},){
			
				@array = ("-","-","-","-","-","-","-","-","-");
				$array_ref = \@array;
				$power-> write_row($rowP, 4, $array_ref, $format_data);

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
									$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "* Toggle_Test*")', $format_data);
									$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "* Contact_Test*")', $format_data);
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
  				open(SourceFile, "<digital/$Mult_file")||open(SourceFile, "<digital/1%$Mult_file");
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
				}
			elsif($testorder{$Mult_file} eq "untest-bscan"){
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_NC);}					#Coverage
				$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  		## Excel ##
				$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);  				## Excel ##
				$rowU++;
				}
			elsif($testorder{$Mult_file} eq "tested-bscan"){
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_NC);}					#Coverage
				$untest-> write($rowU, 0, $Mult_file, $format_data);							## Excel ##
				$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_anno1);  	## Excel ##
				$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);  				## Excel ##
				$rowU++;
				}
		}
	########################################################################################################################################
		}}
	}
	#print $testorder{$Mult_file},$foundTO,"\n";
	################ reservation ###########################################################################################################
	if((not exists ($testorder{$device}) or not exists ($testorder{$Mult_file})) and $foundTO == 0)
	{
		#print $foundTO,"--5--","\n";
		print "			NO Test Found		$device\n"; 
			$coverage-> write($rowC, 5, 'N', $format_NC);			#Coverage
			$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
			$untest-> write($rowU, 1, "NO test item found in TestOrder.", $format_anno);  ## Excel ##
			$untest-> write($rowU, 2, "Check TJ/SP testing.", $format_anno);  ## Excel ##
			$rowU++;
		goto Next_Dev;
		}
   ########################################################################################################################################
#print $testorder{$device},"\n";
Next_Dev:
}

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

###########################################################################################################################################


$bom_coverage_report->close();

print  "\n  >>> Completed ...\n";

END_Prog:

$end_time = time();
$duration = $end_time - $start_time;
printf "\n  runtime: %.4f Sec\n", $duration;

print "\n";
system 'pause';
exit;

