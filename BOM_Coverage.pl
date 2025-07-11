#!/usr/bin/perl
print "\n";
print "*******************************************************************************\n";
print "  Bom Coverage ckecking tool for 3070 <v7.99>\n";
print "  Author: Noon Chen\n";
print "  A Professional Tool for Test.\n";
print "  ",scalar localtime;
print "\n*******************************************************************************\n";
print "\n";

##########################################################################################

use strict;
use warnings;
use Term::ReadKey;
use Time::HiRes qw(time);
use List::Util 'uniq';
use Term::ANSIColor;

our ($Ccode, $bom, $board, $length, $len, $OP, $IC, $pos, $DigPos, $length_DigPin );
our ($dev, $v, $i, $value, $versions, $family, $length_SNail, $NailPos );
our (@array, @testname, @param, @DigPin, @BscanNail, @test_item, %test_nodes );
our ($array_ref, $bdg, $nodes, $fileF, $line, $lineTF, $lineDig, $UTline, $content );
our ($device1, $BscanPin, $node_nam, $Mult_file, $node );
our ($DioNom, $DioHiL, $DioLoL);


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


print "  please specify BOM list file: ";
	$bom=<STDIN>;
	chomp $bom;

############################ Excel #######################################################
use Excel::Writer::XLSX;
use Cwd;
my $currdir = getcwd;

(my $sec,my $min,my $hour,my $mday,my $mon,my $year,my $wday,my $yday,my $isdst) = localtime(time);
#print $hour."-".$min."\n";
#$currdir = `pwd`; chomp $currdir;

$board = substr($currdir,rindex($currdir,"\/")+1);
my $bom_coverage_report = Excel::Writer::XLSX->new($board.'-BOM_Coverage'."-".$hour.$min.$sec.'.xlsx');
my $summary = $bom_coverage_report-> add_worksheet('Summary');
my $coverage = $bom_coverage_report-> add_worksheet('Coverage');
my $tested = $bom_coverage_report-> add_worksheet('Tested');
my $untest = $bom_coverage_report-> add_worksheet('Untest');
my $limited = $bom_coverage_report-> add_worksheet('LimitTest');
my $power = $bom_coverage_report-> add_worksheet('PowerTest');
my $short_thres = $bom_coverage_report-> add_worksheet('Shorts_Setting');

$coverage-> freeze_panes(1,1);			#冻结行、列
$tested-> freeze_panes(1,1);			#冻结行、列
$untest-> freeze_panes(1,0);			#冻结行、列
$limited-> freeze_panes(1,0);			#冻结行、列
$power-> freeze_panes(1,1);				#冻结行、列
$short_thres-> freeze_panes(1,0);		#冻结行、列

$summary-> set_column(0,2,20);			#设置列宽
$coverage-> set_column('A:G',21);		#设置列宽
$tested-> set_column('A:F',20);			#设置列宽
$tested-> set_column('G:G',40);			#设置列宽
$untest-> set_column(0,1,20);			#设置列宽
$untest-> set_column(1,2,40);			#设置列宽
$limited-> set_column(0,1,20);			#设置列宽
$limited-> set_column(1,1,30);			#设置列宽
$power-> set_column(0,1,15);			#设置列宽
$power-> set_column(1,3,30);			#设置列宽
$power-> set_column(4,12,15);			#设置列宽
$short_thres-> set_column(0,1,50);		#设置列宽
$short_thres-> set_column(2,3,20);		#设置列宽

$summary-> activate();					#设置初始可见
$bom_coverage_report->set_size(1680, 1180);	#设置初始窗口尺寸

#新建一个格式
our $format_item = $bom_coverage_report-> add_format(bold=>1, align=>'center', valign=>'vcenter', border=>1, size=>12, bg_color=>'#41D7E8');	# cyan
our $format_head = $bom_coverage_report-> add_format(bold=>1, valign=>'vcenter', border=>1, size=>12, bg_color=>'#6DE841');	# lime
our $format_data = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, text_wrap=>1);
our $format_GND  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'gray');	# gray
our $format_NC   = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'silver');	# silver
our $format_VCC  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'#E87F41');	# orange
our $format_togg = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'#6DE841', text_wrap=>1);	# lime
our $format_pin  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'#2CFF89', text_wrap=>1);	# green
our $format_anno = $bom_coverage_report-> add_format(align=>'left', valign=>'vcenter', border=>1, text_wrap=>1);
our $format_anno1 = $bom_coverage_report-> add_format(align=>'left', valign=>'vcenter', border=>1, text_wrap=>1, bg_color=>'#F4E66A');	# yellow
our $format_PCT  = $bom_coverage_report-> add_format(align=>'center', border=>1, num_format=> '10');
our $format_STP  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, bg_color=>'#F4E66A');	# yellow
our $format_hylk = $bom_coverage_report-> add_format(color=>'blue', align=>'center', valign=>'vcenter', border=>1, underline=>1);
our $format_FPY  = $bom_coverage_report-> add_format(align=>'center', valign=>'vcenter', border=>1, num_format=> '10');

$tested-> write("A1", '<Items>', $format_head);
$tested-> write("B1", '<TYPE>', $format_head);
$tested-> write("C1", '<Board value>', $format_head);
$tested-> write("D1", '<Nominal>', $format_head);
$tested-> write("E1", '<HiLimit>', $format_head);
$tested-> write("F1", '<LoLimit>', $format_head);
$tested-> write("G1", '<Comments>', $format_head);

$untest-> write("A1", '<Items>', $format_head);
$untest-> write("B1", '<Justification>', $format_head);
$untest-> write("C1", '<Comments>', $format_head);

$limited-> write("A1", '<Items>', $format_head);
$limited-> write("B1", '<Comments>', $format_head);

$power-> write("A1", '<Items>', $format_head);
$power-> write("B1", '<TestOrder>', $format_head);
$power-> write("C1", '<TestPlan>', $format_head);
$power-> write("D1", '<Family> , <Test Items>', $format_head);
$power-> write("E1", '<Total Pin>', $format_item);
$power-> write("F1", '<Power Pin>', $format_VCC);
$power-> write("G1", '<GND Pin>', $format_GND);
$power-> write("H1", '<Toggle Test Pin>', $format_togg);
$power-> write("I1", '<Pin Test>', $format_pin);
$power-> write("J1", '<NC Pin>', $format_NC);
$power-> write("K1", '<Untest Pin>', $format_data);
$power-> write("L1", '<Toggle Coverage>', $format_togg);
$power-> write("M1", '<Pin Coverage>', $format_pin);

#$short_thres-> write("A1", 'Fixture TP', $format_head);
#$short_thres-> write("B1", 'Nodes', $format_head);
$short_thres-> write("C1", 'Threshold', $format_head);
$short_thres-> write("D1", 'Delay', $format_head);

$summary-> write("A1", 'Test Items', $format_head);
$summary-> write("B1", 'Quantity', $format_head);
$summary-> write("C1", 'Percentage', $format_head);

$summary-> write("A2", 'Tested', $format_item);
$summary-> write("A3", 'Untest', $format_item);
$summary-> write("A4", 'LimitTest', $format_item);
$summary-> write("A5", 'Power-Tested', $format_item);
$summary-> write("A6", 'Power-UnTest', $format_item);
$summary-> write("A7", 'Node accessibility', $format_item);

$summary-> write("B2", '=COUNTA(Tested!A2:A9999)', $format_data);
$summary-> write("B3", '=COUNTA(Untest!A2:A9999)', $format_data);
$summary-> write("B4", '=COUNTA(LimitTest!A2:A9999)', $format_data);
$summary-> write("B5", '=COUNTIF(PowerTest!C2:C99999,"Tested - *")', $format_data);
$summary-> write("B6", '=SUM(COUNTIF(PowerTest!C2:C99999,{"Unidentified *","skipped -*"}))', $format_data);
$summary-> write("B7", '=COUNTA(Shorts_Setting!B2:B9999)-COUNTIF(Shorts_Setting!C2:C9999,"!*")', $format_data);

$summary-> write_formula("C2", "=(B2/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$summary-> write_formula("C3", "=(B3/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$summary-> write_formula("C4", "=(B4/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$summary-> write_formula("C5", "=(B5/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$summary-> write_formula("C6", "=(B6/(B2+B3+B4+B5+B6))", $format_PCT);  #输出Percentage
$summary-> write_formula("C7", "=(B7/COUNTA(Shorts_Setting!A2:A9999))", $format_PCT);  #输出Percentage

$summary-> write("A21", $currdir);
$summary-> write("A22", '* please update JTAG/Compliance pin coverage manually.');
$summary-> write("A23", '* please scrutinize digital pin coverage.');

$tested-> write("I2", 'Type', $format_item);
$tested-> write("I3", 'Count', $format_item);
$tested-> write("J2", 'Resistor', $format_head);
$tested-> write("K2", 'Capacitor', $format_head);
$tested-> write("L2", 'Inductor', $format_head);
$tested-> write("M2", 'Jumper', $format_head);
$tested-> write("N2", 'Diode', $format_head);
$tested-> write("O2", 'Zener', $format_head);
$tested-> write_formula("J3", '=COUNTIF(B2:B99999,"Resistor")', $format_data);
$tested-> write_formula("K3", '=COUNTIF(B2:B99999,"Capacitor")', $format_data);
$tested-> write_formula("L3", '=COUNTIF(B2:B99999,"Inductor")', $format_data);
$tested-> write_formula("M3", '=COUNTIF(B2:B99999,"Jumper")', $format_data);
$tested-> write_formula("N3", '=COUNTIF(B2:B99999,"Diode")', $format_data);
$tested-> write_formula("O3", '=COUNTIF(B2:B99999,"Zener")', $format_data);

$coverage-> write("A1", ' Device', $format_head);
$coverage-> write("B1", ' L, C, R, D, Z, J test', $format_head);
$coverage-> write("C1", ' Digital logic test', $format_head);
$coverage-> write("D1", ' Analog function test', $format_head);
$coverage-> write("E1", ' Bscan test', $format_head);
$coverage-> write("F1", ' No coverage', $format_head);
$coverage-> write("G1", ' "-" for None', $format_anno);
$coverage-> write("G2", ' "V" for covered', $format_anno);
$coverage-> write("G3", ' "N" for Not covered', $format_anno);
$coverage-> write("G4", ' "L" for paralled', $format_anno);
$coverage-> write("G5", ' "K" for skipped in testplan', $format_anno);

$coverage->conditional_formatting('B2:F99999',
{
	type     => 'text',
	criteria => 'containing',
	value    => 'V',
	format   => $format_togg,
	});

$coverage->conditional_formatting('B2:F99999',
{
	type     => 'text',
	criteria => 'containing',
	value    => 'L',
	format   => $format_pin,
	});

$coverage->conditional_formatting('B2:F99999',
{
	type     => 'text',
	criteria => 'containing',
	value    => 'N',
	format   => $format_NC,
	});

$power-> conditional_formatting('H2:H9999',
{
	type     => 'cell',
	criteria => 'between',
	minimum  => 0.001,
	maximum  => 9999,
	format   => $format_togg,
	});

$power-> conditional_formatting('L2:L9999',
{
	type     => 'cell',
	criteria => 'between',
	minimum  => 0.001,
	maximum  => 9999,
	format   => $format_togg,
	});

$power-> conditional_formatting('I2:I9999',
{
	type     => 'cell',
	criteria => 'between',
	minimum  => 0.001,
	maximum  => 9999,
	format   => $format_pin,
	});

$power-> conditional_formatting('M2:M9999',
{
	type     => 'cell',
	criteria => 'between',
	minimum  => 0.001,
	maximum  => 9999,
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
$summary-> insert_chart('E1', $chart, 1, 1, 1.0, 1.6);

our $rowC = 0;
our $rowT = 1;
our $rowU = 1;
our $rowL = 1;
our $rowP = 1;
our $length_anno = 8;
our $length_TO = 8;
our $length_TP = 8;
our $len_ver = 0;


my $start_time = time();
##################### loading bom to array ###############################################
print "  Gathering all BOM devices...";
our @bom_list = ();

our $number = 0;
open (Export, "> Component_List.txt"); 
open (Import, "< $bom"); 
while(my $array = <Import>)
{
	chomp $array;
	$array =~ s/\s+//g;	   #clear head of line spacing
	next if ($array eq "");
	if ($array =~ "\," and $array !~ "\-")
	{
		our @list = split(/,/, $array);
		#print scalar@list."\n";
		for (our $num = 0; $num < scalar@list; $num++)
		{
			$list[$num] =~ s/(^\s+|\s+$)//g;
			printf Export $list[$num]."\n";
			#print $list[$num]."\n";
			$dev = lc($list[$num]);
			our $bom_list = push(@bom_list, $dev);
			$number++;
			}
		}
	elsif ($array =~ "\-")
	{
		our @fields = split(/,/, $array);
		# print " * ".$array."\n";
		for (our $num = 0; $num < scalar@fields; $num++)
		{
			#print scalar@fields."\n";
			#print "	".$fields[$num]."\n";
			if ($fields[$num] =~ "\-")
			{
				my $string = "";
				my $suffix = "";
				my $begin = "";
				my $final = "";
				my $pattern = "";  # C for character, D for digit	
				# print "	".$fields[$num]."\n";
				my @Comps = split(/-/, $fields[$num]);
				my @Comp = split(/([a-z]+)/i, $Comps[0]);
	
				if ($Comp[1] =~ /([a-z]+)/i and scalar@Comp == 3)
					{$pattern = "CD"; $begin = $Comp[scalar@Comp - 1];}  # begin
	
				if ($Comp[1] =~ /([a-z]+)/i and scalar@Comp == 4)
					{$pattern = "CDC"; $begin = $Comp[scalar@Comp - 2]; $suffix = $Comp[scalar@Comp - 1];}  # begin
	
				if ($Comp[1] =~ /([a-z]+)/i and scalar@Comp == 5)
					{$pattern = "CDCD"; $begin = $Comp[scalar@Comp - 1];}  # begin
	
				@Comp = split(/([a-z]+)/i, $Comps[1]);
				$final = $Comp[scalar@Comp - 1];  # final
				#print $final."\n";
				if ($Comp[1] =~ /([a-z]+)/i and scalar@Comp == 4)
					{$final = $Comp[scalar@Comp - 2];}  # final
				
				# @fields(,) > @Comps(-) > @Comp(c)
				#-------------------------------------------------------------------------
				if ($pattern eq "CD")
				{
					#print "CD\n";
					for (my $num = $begin; $num < $final+1; $num++)
					{
						$string = substr($Comps[1],0,length($Comps[0])-length($begin)).$num;
						if(length($Comps[0]) - length($string) == 1){$string = substr($Comps[0],0,length($Comps[0])-length($begin))."0".$num;}
						if(length($Comps[0]) - length($string) == 2){$string = substr($Comps[0],0,length($Comps[0])-length($begin))."00".$num;}
						if(length($Comps[0]) - length($string) == 3){$string = substr($Comps[0],0,length($Comps[0])-length($begin))."000".$num;}
						if(length($Comps[0]) - length($string) == 4){$string = substr($Comps[0],0,length($Comps[0])-length($begin))."0000".$num;}
						
						$string =~ s/(^\s+|\s+$)//g;
						printf Export $string."\n";
						#print $string."\n";
						$dev = lc($string);
						our $bom_list = push(@bom_list, $dev);
						$number++;
						}
					}
	
				#-------------------------------------------------------------------------
				if ($pattern eq "CDC")
				{
					#print "CDC\n";
					for (my $num = $begin; $num < $final+1; $num++)
					{
						$string = substr($Comps[0],0,index($Comps[0],$begin)).$num.$suffix;
						if(length($Comps[0]) - length($string) == 1){$string = substr($Comps[0],0,index($Comps[0],$begin))."0".$num.$suffix;}
						if(length($Comps[0]) - length($string) == 2){$string = substr($Comps[0],0,index($Comps[0],$begin))."00".$num.$suffix;}
						if(length($Comps[0]) - length($string) == 3){$string = substr($Comps[0],0,index($Comps[0],$begin))."000".$num.$suffix;}
						if(length($Comps[0]) - length($string) == 4){$string = substr($Comps[0],0,index($Comps[0],$begin))."0000".$num.$suffix;}
						
						$string =~ s/(^\s+|\s+$)//g;
						printf Export $string."\n";
						#print $string."\n";
						$dev = lc($string);
						our $bom_list = push(@bom_list, $dev);
						$number++;
						}
					}
	
				#-------------------------------------------------------------------------
				if ($pattern eq "CDCD")
				{
					#print "CDCD\n";
					for (my $num = $begin; $num < $final+1; $num++)
					{
						$string = substr($Comps[0],0,length($Comps[0])-length($begin)).$num;
						if(length($Comps[0]) - length($string) == 1){$string = substr($Comps[0],0,length($Comps[0])-length($begin))."0".$num;}
						if(length($Comps[0]) - length($string) == 2){$string = substr($Comps[0],0,length($Comps[0])-length($begin))."00".$num;}
						if(length($Comps[0]) - length($string) == 3){$string = substr($Comps[0],0,length($Comps[0])-length($begin))."000".$num;}
						if(length($Comps[0]) - length($string) == 4){$string = substr($Comps[0],0,length($Comps[0])-length($begin))."0000".$num;}
						
						$string =~ s/(^\s+|\s+$)//g;
						printf Export $string."\n";
						#print $string."\n";
						$dev = lc($string);
						our $bom_list = push(@bom_list, $dev);
						$number++;
						}
					}
				}
			else
			{
				$fields[$num] =~ s/(^\s+|\s+$)//g;
				printf Export $fields[$num]."\n";
				#print $fields[$num]."\n";
				$dev = lc($fields[$num]);
				our $bom_list = push(@bom_list, $dev);
				$number++;
				}
			}
		}
	else
	{
		$array =~ s/(^\s+|\s+$)//g;
		print Export $array,"\n";
		#print $array,"\n";
		$dev = lc($array);
		our $bom_list = push(@bom_list, $dev);
		$number++;
		}
	}

close Import;
close Export;

print "[DONE]\n";
print "\n	BOM count: ".$number."\n";
$summary-> write("A11", "BOM count: ".$number);

@bom_list = sort @bom_list;
@bom_list = uniq @bom_list;
$length = scalar @bom_list;
print "	valid BOM: ", $length,"\n";
$summary-> write("A12", "valid BOM: ".$length);

##################### loading BOMvalue to hash ###########################################
my %bom_value = ();
open (Board, "< board") or warn "\t!!! Failed to open 'board' file: $!.\n";
	if ($! eq "No such file or directory"){print "\n\t>>> program exiting ...\n"; <STDIN>; exit;}
while($line = <Board>)
{
	chomp ($line);
	$line =~ s/(^\s+|\s+$)//g;
	next if ($line eq '');

	if($line eq "CAPACITOR")
	{
		while($line = <Board>)
		{
			chomp ($line);
			$line =~ s/(^\s+|\s+$)//g;
			last if ($line eq '');
			my @line = ();
			@line = split('\ ',$line);
			last if (scalar @line < 5);
			#print $line[0],"-",$line[1],"\n";
			$bom_value{$line[0]} = $line[1];
			}
		}

	if($line eq "JUMPER")
	{
		while($line = <Board>)
		{
			chomp ($line);
			$line =~ s/(^\s+|\s+$)//g;
			last if ($line eq '');
			my @line = ();
			@line = split('\ ', $line, 3);
			if($line =~ " PN"){@line = split('\ ', $line, 4);}
			next unless (scalar @line > 2);
			next if (scalar @line == 3 and uc(substr($line[2],1,3)) ne "RES");
			next if (scalar @line == 4 and uc(substr($line[3],1,3)) ne "RES");
			if(scalar @line == 3 and uc(substr($line[2],1,3)) eq "RES")
			{
				my @line1 = split( substr($line[2],4,1), $line[2]);
				#print $line[0],"-",$line[2],"\n";
				if (substr($line1[1],-1,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = $line1[1]; next;}
				elsif (substr($line1[2],-1,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = $line1[2]; next;}
				elsif (substr($line1[3],-1,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = $line1[3]; next;}

				elsif (uc(substr($line1[1],-1,1)) eq "K" and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-1); next;}
				elsif (uc(substr($line1[2],-1,1)) eq "K" and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-1); next;}
				elsif (uc(substr($line1[3],-1,1)) eq "K" and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-1); next;}

				elsif (uc(substr($line1[1],-1,1)) eq "M" and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-1); next;}
				elsif (uc(substr($line1[2],-1,1)) eq "M" and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-1); next;}
				elsif (uc(substr($line1[3],-1,1)) eq "M" and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-1); next;}
				
				elsif (uc(substr($line1[1],-3,3)) eq "OHM"){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-3); next;}
				elsif (uc(substr($line1[2],-3,3)) eq "OHM"){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-3); next;}
				elsif (uc(substr($line1[3],-3,3)) eq "OHM"){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-3); next;}
				else {print "\t\t!!! failed to read value for: $line[0]\n";}
				}
			if(scalar @line == 4 and uc(substr($line[3],1,3)) eq "RES")
			{
				my @line1 = split( substr($line[3],4,1), $line[3]);
				#print $line[0],"-",$line[3],"\n";
				if (substr($line1[1],-1,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = $line1[1]; next;}
				elsif (substr($line1[2],-1,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = $line1[2]; next;}
				elsif (substr($line1[3],-1,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = $line1[3]; next;}

				elsif (uc(substr($line1[1],-1,1)) eq "K" and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-1); next;}
				elsif (uc(substr($line1[2],-1,1)) eq "K" and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-1); next;}
				elsif (uc(substr($line1[3],-1,1)) eq "K" and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-1); next;}

				elsif (uc(substr($line1[1],-1,1)) eq "M" and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-1); next;}
				elsif (uc(substr($line1[2],-1,1)) eq "M" and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-1); next;}
				elsif (uc(substr($line1[3],-1,1)) eq "M" and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-1); next;}
				
				elsif (uc(substr($line1[1],-3,3)) eq "OHM"){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-3); next;}
				elsif (uc(substr($line1[2],-3,3)) eq "OHM"){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-3); next;}
				elsif (uc(substr($line1[3],-3,3)) eq "OHM"){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-3); next;}
				else {print "\t\t!!! failed to read value for: $line[0]\n";}
				}
			}
		}

	if($line eq "PIN LIBRARY")
	{
		while($line = <Board>)
		{
			chomp ($line);
			$line =~ s/(^\s+|\s+$)//g;
			last if ($line eq '');
			my @line = ();
			@line = split('\ ', $line, 3);
			if($line =~ " NP"){@line = split('\ ', $line, 4);}
			next if (scalar @line < 3);
			next if (scalar @line == 3 and uc(substr($line[2],1,3)) !~ /(CAP|RES)/);
			next if (scalar @line == 4 and uc(substr($line[3],1,3)) !~ /(CAP|RES)/);
			if(scalar @line == 3 and uc(substr($line[2],1,3)) eq "CAP")
			{
				my @line1 = split( substr($line[2],4,1), $line[2]);
				#print $line[0],"-",$line[2],"\n";
				if (uc(substr($line1[1],-1,1)) =~ /^(M|U|N|P)$/ and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = lc($line1[1]); next;}
				elsif (uc(substr($line1[2],-1,1)) =~ /^(M|U|N|P)$/ and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = lc($line1[2]); next;}
				elsif (uc(substr($line1[3],-1,1)) =~ /^(M|U|N|P)$/ and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = lc($line1[3]); next;}
				
				elsif (uc(substr($line1[1],-2,2)) =~ /^(MF|UF|NF|PF)$/ and substr($line1[1],-3,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = lc(substr($line1[1],0,-1)); next;}
				elsif (uc(substr($line1[2],-2,2)) =~ /^(MF|UF|NF|PF)$/ and substr($line1[2],-3,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = lc(substr($line1[2],0,-1)); next;}
				elsif (uc(substr($line1[3],-2,2)) =~ /^(MF|UF|NF|PF)$/ and substr($line1[3],-3,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = lc(substr($line1[3],0,-1)); next;}
				else {print "\t\t!!! failed to read value for: $line[0]\n";}
				}
			if(scalar @line == 3 and uc(substr($line[2],1,3)) eq "RES")
			{
				my @line1 = split( substr($line[2],4,1), $line[2]);
				#print $line[0],"-",$line[2],"\n";
				if (substr($line1[1],-1,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = $line1[1]; next;}
				elsif (substr($line1[2],-1,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = $line1[2]; next;}
				elsif (substr($line1[3],-1,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = $line1[3]; next;}

				elsif (uc(substr($line1[1],-1,1)) eq "K" and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-1); next;}
				elsif (uc(substr($line1[2],-1,1)) eq "K" and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-1); next;}
				elsif (uc(substr($line1[3],-1,1)) eq "K" and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-1); next;}

				elsif (uc(substr($line1[1],-1,1)) eq "M" and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-1); next;}
				elsif (uc(substr($line1[2],-1,1)) eq "M" and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-1); next;}
				elsif (uc(substr($line1[3],-1,1)) eq "M" and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-1); next;}
				
				elsif (uc(substr($line1[1],-3,3)) eq "OHM"){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-3); next;}
				elsif (uc(substr($line1[2],-3,3)) eq "OHM"){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-3); next;}
				elsif (uc(substr($line1[3],-3,3)) eq "OHM"){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-3); next;}
				else {print "\t\t!!! failed to read value for: $line[0]\n";}
				}
			if(scalar @line == 4 and uc(substr($line[3],1,3)) eq "CAP")
			{
				my @line1 = split( substr($line[3],4,1), $line[3]);
				#print $line[0],"-",$line[2],"\n";
				if (uc(substr($line1[1],-1,1)) =~ /^(M|U|N|P)$/ and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = lc($line1[1]); next;}
				elsif (uc(substr($line1[2],-1,1)) =~ /^(M|U|N|P)$/ and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = lc($line1[2]); next;}
				elsif (uc(substr($line1[3],-1,1)) =~ /^(M|U|N|P)$/ and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = lc($line1[3]); next;}
				
				elsif (uc(substr($line1[1],-2,2)) =~ /^(MF|UF|NF|PF)$/ and substr($line1[1],-3,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = lc(substr($line1[1],0,-1)); next;}
				elsif (uc(substr($line1[2],-2,2)) =~ /^(MF|UF|NF|PF)$/ and substr($line1[2],-3,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = lc(substr($line1[2],0,-1)); next;}
				elsif (uc(substr($line1[3],-2,2)) =~ /^(MF|UF|NF|PF)$/ and substr($line1[3],-3,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = lc(substr($line1[3],0,-1)); next;}
				else {print "\t\t!!! failed to read value for: $line[0]\n";}
				}
			if(scalar @line == 4 and uc(substr($line[3],1,3)) eq "RES")
			{
				my @line1 = split( substr($line[3],4,1), $line[3]);
				#print $line[0],"-",$line[3],"\n";
				if (substr($line1[1],-1,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = $line1[1]; next;}
				elsif (substr($line1[2],-1,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = $line1[2]; next;}
				elsif (substr($line1[3],-1,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = $line1[3]; next;}

				elsif (uc(substr($line1[1],-1,1)) eq "K" and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-1); next;}
				elsif (uc(substr($line1[2],-1,1)) eq "K" and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-1); next;}
				elsif (uc(substr($line1[3],-1,1)) eq "K" and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-1); next;}

				elsif (uc(substr($line1[1],-1,1)) eq "M" and substr($line1[1],-2,1) =~ '\d'){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-1); next;}
				elsif (uc(substr($line1[2],-1,1)) eq "M" and substr($line1[2],-2,1) =~ '\d'){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-1); next;}
				elsif (uc(substr($line1[3],-1,1)) eq "M" and substr($line1[3],-2,1) =~ '\d'){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-1); next;}
				
				elsif (uc(substr($line1[1],-3,3)) eq "OHM"){$line1[1] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[1],0,-3); next;}
				elsif (uc(substr($line1[2],-3,3)) eq "OHM"){$line1[2] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[2],0,-3); next;}
				elsif (uc(substr($line1[3],-3,3)) eq "OHM"){$line1[3] =~ s/\s+//g; $bom_value{$line[0]} = substr($line1[3],0,-3); next;}
				else {print "\t\t!!! failed to read value for: $line[0]\n";}
				}
			}
		}

	if($line eq "RESISTOR")
	{
		while($line = <Board>)
		{
			chomp ($line);
			$line =~ s/(^\s+|\s+$)//g;
			last if ($line eq '');
			my @line = ();
			@line = split('\ ',$line);
			last if (scalar @line < 5);
			#print $line[0],"-",$line[1],"\n";
			$bom_value{$line[0]} = $line[1];
			}
		}
	}
close Board;

my @keysBOM = keys %bom_value;
my $sizeBOM = @keysBOM;
print "\tBOM value: ", $sizeBOM,"\n";
$summary-> write("A13", "BOM value: ".$sizeBOM);

# print "CF036 value is: $bom_value{'CF036'} \n";

# while (my ($key, $value) = each %bom_value) {
#     print "$key => $value\n";
# }

# foreach my $key (sort keys %bom_value) {
#     print "$key => $bom_value{$key}\n";
# }


##################### loading BDG to hash ################################################
my %bdg_list = ();

my $BDG_File = "./bdg_data/dig_inc_ver_fau.dat";
if(-e $BDG_File)
{
	open (BDGFile, "< ./bdg_data/dig_inc_ver_fau.dat");
	while($bdg = <BDGFile>)
	{
		chomp ($bdg);
		$bdg =~ s/(^\s+|\s+$)//g;
		next if ($bdg eq '');
		my @bdg = ();
		@bdg = split('\"',$bdg);

		if($bdg[0] =~ "Verifying faults on")
		{
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
else{warn "\t!!! Failed to open './bdg_data/dig_inc_ver_fau.dat': $!.\n";}
my @keysBDG = keys %bdg_list;
my $sizeBDG = @keysBDG;
print "	BDG: ", $sizeBDG,"\n";
$summary-> write("A14", "BDG: ".$sizeBDG);


##################### loading pins to hash ###############################################
my %hash_pin = ();
open (Pin, "< pins") or open (Pin, "< 1%pins") or warn "\t!!! Failed to open 'pins' file: $!.\n";
while($nodes = <Pin>)
{
	chomp ($nodes);
	$nodes =~ s/(^\s+|\s+$)//g;
	next if ($nodes =~ "^\!");
	next if ($nodes eq "");
	my @nodes = ();
	@nodes = split('\"',$nodes);
	if ($nodes[1] =~ "\%"){$nodes[1] = substr($nodes[1],2);}
	
	if (substr($nodes[0], 0, 5) eq "nodes")
	{
		$hash_pin{$nodes[1]} = 1;
		#print $nodes[0],"-- ",$nodes[1]."\n";
		}
	}
close Pin;

my @keys = keys %hash_pin;
my $size = @keys;
print "	Pins: ".$size."\n";
$summary-> write("A15", "Pins: ".$size);


##################### loading testorder to hash ##########################################
my %testorder = ();
my @versions = ();
open (TO, "< testorder") or warn "\t!!! Failed to open 'testorder' file: $!.\n";
	if ($! eq "No such file or directory"){print "\n\t>>> program exiting ...\n"; <STDIN>; exit;}
while($dev = <TO>)
{
	chomp($dev);
	$dev =~ s/(^\s+|\s+$)//g;
	next unless ($dev =~ "^test" or $dev =~ "^skip");
	next if ($dev =~ "scan disable");
	$value = "";
	my @dev = ();
	@dev = split('\"',$dev);
	$dev[0] =~ s/(^\s+|\s+$)//g;
	$dev[1] =~ s/(^\s+|\s+$)//g;
	if (scalar @dev > 2){$dev[2] =~ s/(^\s+|\s+$)//g; $dev[2] =~ s/( +)//g;}
	else{$dev[2] = "";}

	if ($dev[2] ne "version")
	{
		if ($dev[0] eq "test resistor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-res'}
		if ($dev[0] eq "test capacitor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-cap'}
		if ($dev[0] eq "test jumper" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-jmp'}
		if ($dev[0] eq "test diode" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-dio'}
		if ($dev[0] eq "test zener" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-zen'}
		if ($dev[0] eq "test inductor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-ind'}
		if ($dev[0] eq "test fuse" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'tested-fuse'}

		if ($dev[0] eq "skip resistor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'skipped-res'}
		if ($dev[0] eq "skip capacitor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'skipped-cap'}
		if ($dev[0] eq "skip jumper" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'skipped-jmp'}
		if ($dev[0] eq "skip diode" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'skipped-dio'}
		if ($dev[0] eq "skip zener" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'skipped-zen'}
		if ($dev[0] eq "skip inductor" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'skipped-ind'}
		if ($dev[0] eq "skip fuse" and ($dev[2] eq "" or $dev[2] ne "nulltest")){$value = 'skipped-fuse'}

		if ($dev[0] =~ "resistor" and $dev[2] eq ";nulltest"){$value = 'untest-res'}
		if ($dev[0] =~ "capacitor" and $dev[2] eq ";nulltest"){$value = 'untest-cap'}
		if ($dev[0] =~ "jumper" and $dev[2] eq ";nulltest"){$value = 'untest-jmp'}
		if ($dev[0] =~ "diode" and $dev[2] eq ";nulltest"){$value = 'untest-dio'}
		if ($dev[0] =~ "zener" and $dev[2] eq ";nulltest"){$value = 'untest-zen'}
		if ($dev[0] =~ "inductor" and $dev[2] eq ";nulltest"){$value = 'untest-ind'}
		if ($dev[0] =~ "fuse" and $dev[2] eq ";nulltest"){$value = 'untest-fuse'}

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
		#print $dev[1],"	--	",$value,"\n";
		}
	elsif ($dev[2] eq "version")
	{
		if ($dev[0] eq "test resistor"){$value = 'tested-res+ver'}
		if ($dev[0] eq "test capacitor"){$value = 'tested-cap+ver'}
		if ($dev[0] eq "test jumper"){$value = 'tested-jmp+ver'}
		if ($dev[0] eq "test diode"){$value = 'tested-dio+ver'}
		if ($dev[0] eq "test zener"){$value = 'tested-zen+ver'}
		if ($dev[0] eq "test inductor"){$value = 'tested-ind+ver'}
		if ($dev[0] eq "test fuse"){$value = 'tested-fuse+ver'}
		
		if ($dev[0] eq "skip resistor"){$value = 'untest-res+ver'}
		if ($dev[0] eq "skip capacitor"){$value = 'untest-cap+ver'}
		if ($dev[0] eq "skip jumper"){$value = 'untest-jmp+ver'}
		if ($dev[0] eq "skip diode"){$value = 'untest-dio+ver'}
		if ($dev[0] eq "skip zener"){$value = 'untest-zen+ver'}
		if ($dev[0] eq "skip inductor"){$value = 'untest-ind+ver'}
		if ($dev[0] eq "skip fuse"){$value = 'untest-fuse+ver'}
		
		if ($dev[0] eq "test mixed"){$value = 'tested-mix+ver'}
		if ($dev[0] eq "test digital"){$value = 'tested-dig+ver'}
		if ($dev[0] eq "test analog powered"){$value = 'tested-pwr+ver'}
		if ($dev[0] eq "test scan connect"){$value = 'tested-bscan+ver'}
		
		if ($dev[0] eq "skip mixed"){$value = 'untest-mix+ver'}
		if ($dev[0] eq "skip digital"){$value = 'untest-dig+ver'}
		if ($dev[0] eq "skip analog powered"){$value = 'untest-pwr+ver'}
		if ($dev[0] eq "skip scan connect"){$value = 'untest-bscan+ver'}
		
		$versions = push(@versions, $dev[3]);
		$testorder{$dev[3]."+".$dev[1]} = $value;
		#print $dev[3]."+".$dev[1],"	--	",$value,"\n";
		}
	}
close TO;

my @keysTO = keys %testorder;
my $sizeTO = @keysTO;
print "	testorder: ".$sizeTO."\n";
$summary-> write("A16", "testorder: ".$sizeTO);

@versions = uniq @versions;
$len_ver = scalar @versions;
print "	$len_ver versions: @versions\n";
$summary-> write("A17", "$len_ver versions: @versions");


##################### loading testplan to hash ###########################################
our %testplan = ();
my $learn = 0;
open (TP, "< testplan") or warn "\t!!! Failed to open 'testplan' file: $!.\n";
	if ($! eq "No such file or directory"){print "\n\t>>> program exiting ...\n"; <STDIN>; exit;}
while($dev = <TP>)
{
	my @dev = ();
	chomp ($dev);
	$dev =~ s/(^\s+|\s+$)//g;
	if ($dev =~ "learn capacitance on"){$learn = 1;}
	if ($dev =~ "learn capacitance off"){$learn = 0;}
	next if ($learn == 1);
	next unless ($dev =~ "test" and $dev =~ "\"");
	
	$value = '';
	@dev = split('\"',$dev);
	$dev[0] =~ s/ *//g;
	next unless ($dev[0] eq "test" or ($dev[0] =~ "test" and substr($dev[0],0,1) eq "\!"));

	$dev[0] =~ s/(^\s+|\s+$)//g;
	$dev[1] =~ s/(^\s+|\s+$)//g;
	if (scalar @dev > 2){$dev[2] =~ s/(^\s+|\s+$)//g;}
	else{$dev[2] = "";}
	
	if ($dev[0] eq "test")
	{
		if($dev[1] =~ 'mixed/'){$dev[1] = substr($dev[1],6)}
		if($dev[1] =~ 'analog/'){$dev[1] = substr($dev[1],7)}
		if($dev[1] =~ 'digital/'){$dev[1] = substr($dev[1],8)}
		if($dev[2] =~ "\!"){$value = "tested".substr($dev[2],rindex($dev[2],"\!") + 1);}
		if($dev[2] !~ "\!"){$value = "tested";}
		}
	elsif ($dev[0] =~ "test" and substr($dev[0],0,1) eq "\!")
	{
		if($dev[1] =~ 'mixed/'){$dev[1] = substr($dev[1],6)}
		if($dev[1] =~ 'analog/'){$dev[1] = substr($dev[1],7)}
		if($dev[1] =~ 'digital/'){$dev[1] = substr($dev[1],8)}
		if($dev[2] =~ "\!"){$value = "skipped-".substr($dev[2],rindex($dev[2],"\!") + 1);}
		if($dev[2] !~ "\!"){$value = "skipped-";}
		}
	else{$value = "unidentified";}
	$testplan{$dev[1]} = $value;
	#print $dev[1],"	--	",$value."\n";
	#print substr($testplan{$dev[1]},0,6),"\n";
	}
close TP;

my @keysTP = keys %testplan;
my $sizeTP = @keysTP;
print "	testplan: ".$sizeTP."\n";
$summary-> write("A18", "testplan: ".$sizeTP);


############################### shorts threshold statistic ###############################
print  "\n	>>> Analyzing shorts threshold ...\n";

my $delay = "-";
my $thres = "-";
$node = 1;
%test_nodes = ();
my @skip_nodes = ();
our @shorts = ();

open (Thres, "< shorts") or open (Thres, "< 1%shorts") or warn "\t!!! Failed to open 'shorts' file: $!.\n";
if ($! ne "No such file or directory")
{
	while($nodes = <Thres>)
	{
		chomp $nodes;
		$nodes =~ s/^ +//;	   #clear head of line spacing
		$nodes =~ s/( +)/ /g;
		if (substr($nodes,0,9) =~ "threshold")
		{
			$thres = substr($nodes, index($nodes,"threshold")+10);
			if ($nodes =~ "\!"){$thres = substr($nodes, 10, index($nodes,"\!")-10);}
			$thres =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			}
		if (substr($nodes,0,14) =~ "settling delay") 
		{
			#print $nodes,"\n";
			if($nodes =~ "\!"){$delay = substr($nodes, 15, index($nodes,"\!")-15);}
			else{$delay = substr($nodes, 15);}
			#print $delay,"\n";
			#if ($nodes =~ "\!"){$delay = substr($nodes, 10, index($nodes,"\!")-10);}
			$delay =~ s/(^\s+|\s+$)//g;                     #clear all spacing
			}
		if ($nodes =~ "nodes")
		{
			my @node = split('\"', $nodes);
			if(substr($nodes,0,1) eq "!")
			{
				my $node_name = $node[1];
				$node_name =~ s/(^\s+|\s+$)//g;                     #clear all spacing
				#$short_thres-> write($node, 0, $node_name, $format_data);  ## Nodes ##
				#$short_thres-> write($node, 1, substr($nodes, rindex($nodes,"!")), $format_data);  ## Thres ##
				#$short_thres-> write($node, 2, "-", $format_data);  ## Delay ##
				push (@skip_nodes, $node_name."/".substr($nodes, rindex($nodes,"!"))."/"."-")
				}
			elsif(substr($nodes,0,5) eq "nodes")
			{
				#$short_thres-> write($node, 0, $nodes, $format_data);  ## Nodes ##
				#$short_thres-> write($node, 1, $thres, $format_data);  ## Thres ##
				#$short_thres-> write($node, 2, $delay, $format_data);  ## Delay ##
				#push (%test_nodes, $nodes."/".$thres."/".$delay)
				$test_nodes{$node[1]} = $thres."/".$delay;
				}
			#$node++;
			#print $nodes."/".$thres."/".$delay."\n";
			}
		if (substr($nodes,0,5) eq "short") 
		{
			$nodes =~ s/( +)//g;
			#print $nodes,"\n";
			my @nodes = split('\"', uc($nodes));
			push (@shorts, $nodes[1]."/".$nodes[3]);
			}
		}
	}
close Thres;


foreach my $key (sort keys %test_nodes)
{
    #print "$key => $test_nodes{$key}\n";
    @test_item = split("\/", $test_nodes{$key});
    $short_thres-> write($node, 1, $key, $format_data);  ## Nodes ##
    $short_thres-> write($node, 2, $test_item[0], $format_data);  ## Thres ##
    $short_thres-> write($node, 3, $test_item[1], $format_data);  ## Delay ##
	$node++;
	}

foreach my $i (0..@skip_nodes-1)
{
	#print $skip_nodes[$i];
	@test_item = split("\/", $skip_nodes[$i]);
	$short_thres-> write($node, 0, "-", $format_data);
	$short_thres-> write($node, 1, $test_item[0], $format_anno);  ## Nodes ##
	$short_thres-> write($node, 2, $test_item[1], $format_anno);  ## Thres ##
	$short_thres-> write($node, 3, $test_item[2], $format_data);  ## Delay ##
	$node++;
	}

my @keysNode = keys %test_nodes;
my $sizeNode = @keysNode;
print "\ttested shorts: ". @shorts ."\n";
print "\ttested nodes: $sizeNode\n";
$short_thres-> write("B1", 'Tested Nodes: '.$sizeNode, $format_head);


############################### reading fixture.o ########################################
my $Nnum = 0;	#node numbers
my $Wnum = 0;   #Wire numbers
$node = 1;
my @pins =(); my $short_pair = ''; my $BRC = '';
my $BRCbuffer = ''; my @node = ''; my $panel = "";

open (Fixture, "< ./fixture/fixture.o") or warn "\t!!! Failed to open 'fixture.o' file: $!.\n";
open (Report, ">Fixture.txt");
while(my $LIST = <Fixture>)
{
	$LIST =~ s/(^\s+|\s+$)//g;		#clear all non-character symbol
	next if(!$LIST);				#goto next if it's empty
	my @nodes = split('\s+', $LIST);
	last if ($LIST eq "PROTECTED UNIT");
	if ($LIST =~ "END PANEL"){$panel = "True";}

	if($nodes[0] eq "NODE")
	{
		$LIST = <Fixture>;
		$LIST =~ s/(^\s+|\s+$)//g;
		next if(!$LIST);			#goto next if it's empty
		if($LIST ne "PROBES")
		{
			$Nnum++;
			if($nodes[1] =~ '\%'){$nodes[1] = substr($nodes[1], 3, -1)}
			if($nodes[1] =~ '\"'){$nodes[1] = substr($nodes[1], 1, -1)}
			#next if($nodes[1] =~ /(^NC_|_NC$|NONE)/);	#eliminate NC nets
			#print "Probe\#:$Nnum	$nodes[1]\n";
			print Report "	#$Nnum\n";
			print Report "$nodes[1]\n";
			while($LIST = <Fixture>)
			{
				$LIST =~ s/(^\s+|\s+$)//g;
				last if(!$LIST);		#exit loop if it's none-character symbol				
				if($LIST eq "WIRES")
				{
					my @pair = ();
					#print @pair."\n";
					my $BRCnum = 0;	#BRC numbers
					while($LIST = <Fixture>)
					{
						$LIST =~ s/(^\s+|\s+$)//g;	   #clear all non-character symbol
						goto NEXT_NODE if(!$LIST);		#exit loop if it's none-character symbol
						@node = split('\s+', $LIST);
						if($node[0] !~ /(\D+)/) {($BRC) = $node[0] =~ /(\d+)/;
							next if(substr($BRC,-2) =~ /(19|20|39|40|59|60)/);	#eliminate fixed GROUND
							next if(substr($BRC,0,3) =~ /(201|213|111|123|106|118|206|218)/);	#eliminate ASRU/Control card				
							next if($BRCbuffer eq $BRC);
							$Wnum++;
							#print "   Wire\#:$Wnum	",$BRC,"\n";
							print Report $BRC."\n";
							$BRCbuffer = $BRC;
							unshift(@pair, $BRC);
							if($BRCnum > 0)
							{
								#print $BRC,"\n";
								if($BRC < $pair[$BRCnum]){$short_pair = $BRC." to ".$pair[$BRCnum]."\n";}
								if($BRC > $pair[$BRCnum]){$short_pair = $pair[$BRCnum]." to ".$BRC."\n";}
								print Report " -- ".$short_pair;
								push(@pins, $nodes[1]);
								}
							else		#collect pins data
							{
								push(@pins, $nodes[1]);
								}
							$BRCnum++;
							}
						}
					}
				}
			}
		}
	NEXT_NODE:
	}
close Report;
close Fixture;

my @unique_pins = uniq @pins;
@unique_pins = sort @unique_pins;
#my $pins_ref = \@unique_pins;
#$short_thres-> write_col(1, 0, $pins_ref, $format_data);

foreach my $i (0..@unique_pins-1)
{
	#print $unique_pins[$i],"\n";
	if (exists $test_nodes{$unique_pins[$i]})
	{
		$short_thres-> write($node, 0, $unique_pins[$i], $format_data);  ## Nodes ##
		}
	else
	{
		$short_thres-> write($node, 0, $unique_pins[$i], $format_STP);  ## Nodes ##
		}
	
	$node++;
	}
print "\tfixture nodes: ".scalar@unique_pins."\n";
$summary-> write("A10", "fixture nodes: ".scalar @unique_pins);

if (scalar @unique_pins == $sizeNode)
{
	$short_thres-> write("A1", 'Fixture TP: '.scalar @unique_pins, $format_head);
	}
else
{
	$short_thres-> write("A1", 'Fixture TP: '.scalar @unique_pins, $format_anno1);
	}

#-----------------------------------------------------------------------------------------

my $row_ver = 0;
our @Paral = ("tested-res","tested-cap");
our @TOtested = ("tested-res","tested-cap","tested-jmp","tested-dio","tested-zen","tested-ind","tested-fuse");
our @TOtestedver = ("tested-res+ver","tested-cap+ver","tested-jmp+ver","tested-dio+ver","tested-zen+ver","tested-ind+ver","tested-fuse+ver");
our @TOuntest = ("untest-res","untest-cap","untest-jmp","untest-dio","untest-zen","untest-ind","untest-fuse");
our @TOuntestver = ("untest-res+ver","untest-cap+ver","untest-jmp+ver","untest-dio+ver","untest-zen+ver","untest-ind+ver","untest-fuse+ver");
our @TOskipped = ("skipped-res","skipped-cap","skipped-jmp","skipped-dio","skipped-zen","skipped-ind","skipped-fuse");
our @TOparal = ("paral-res","paral-cap","paral-jmp","paral-dio","paral-zen","paral-ind","paral-fuse");

##########################################################################################
print "\n";
foreach my $device (@bom_list)
{
	my $Total_Pin = 0;
	my $Power_Pin = 0;
	my $GND_Pin = 0;
	my $Toggle_Pin = 0;
	my $NC_Pin = 0;
	my $rowP_ori = $rowP;
	my $ibus = ""; my $sbus = ""; my $state = "";
	#-------------------------------------------------------------------------------------
	my $worksheet = 0;
	my $foundTO = 0;
	my $foundTP = 0;
	$device =~ s/(^\s+|\s+$)//g;                     #clear all spacing
	$device = lc($device);

	$rowC = $rowC+1;	my $cover = 0;	my $UNCover = 0;
	$coverage-> write($rowC, 0, $device, $format_data);			#Coverage
	@array = ("-","-","-","-","-");
	$array_ref = \@array;
	$coverage-> write_row($rowC, 1, $array_ref, $format_data);
	print "	Analyzing ", $device, " .....\n";

	#print $testorder{$device},"\n";
	#print $testplan{$device},"\n";
	################ testable device #####################################################
	if(exists($testplan{$device})
	and substr($testplan{$device},0,6) eq "tested"
	and grep{ $_ eq $testorder{$device}} @TOtested)
	{
		$foundTO = 1;
		print "			Common AnaTest		", $device."\n";
		@array = ("-","-","-","-","-");
		$array_ref = \@array;
		$tested-> write_row($rowT, 2, $array_ref, $format_data);
	
		$UNCover = 1;
		$coverage-> write($rowC, 1, 'V', $format_data);			#Coverage
		$row_ver = $rowT;
		
		#$value = read_value(uc($device));
		$value = '-';
		if(exists($bom_value{uc($device)})){$value = $bom_value{uc($device)};} #print "$device value is: $bom_value{uc($device)} \n"; }
		
		open(SourceFile, "<analog/$device") or open(SourceFile, "<analog/1%$device") or warn "\t!!! Failed to open '$device' file: $!.\n";
		if ($! eq "No such file or directory"){$tested-> write($rowT, 6, "TestFile not found.", $format_anno1);}
		else
		{
		while($lineTF = <SourceFile>)							#read parameter
		{
			$len = 0;
			chomp($lineTF);
			$lineTF =~ s/^ +//;                               	#clear head of line spacing
			#print $lineTF,"\n";
			next if (length($lineTF) < 2);
			my @list = split('\ ', $lineTF, 2);
			#print $list[0],"\n";
			#print $list[1],"\n";
			#print $value,"\n";

			if ($list[0] eq "connect" and substr($list[1],0,1) eq "i" and $list[1] !~ " pins "){my @list1 = split('\"', $list[1], 3); $ibus = uc($list1[1]); }
			if ($list[0] eq "connect" and substr($list[1],0,1) eq "s" and $list[1] !~ " pins "){my @list1 = split('\"', $list[1], 3); $sbus = uc($list1[1]); }

			if (grep{ $_ eq $testorder{$device}} @Paral and $lineTF =~ "report parallel devices")
			{
				while($lineTF = <SourceFile>)
				{
					chomp($lineTF);
					$lineTF =~ s/^ +//;
					$lineTF =~ s/( +)/ /g;
					last if ($lineTF =~ "end on failure");
					#print $lineTF,"\n";
					next if (length($lineTF) < 2);
					my @Vlist = split('\ ', $lineTF);
					#print substr($Vlist[2],-2,1),substr($value,-1,1),"\n";

					if(substr($value,-1,1) eq "m" and substr($Vlist[2],-2,1) eq "m") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2))."m";}
					if(substr($value,-1,1) eq "m" and substr($Vlist[2],-2,1) eq "u") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)/1000)."m";}

					if(substr($value,-1,1) eq "u" and substr($Vlist[2],-2,1) eq "u") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2))."u";}
					if(substr($value,-1,1) eq "u" and substr($Vlist[2],-2,1) eq "m") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000)."u";}
					if(substr($value,-1,1) eq "u" and substr($Vlist[2],-2,1) eq "n") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)/1000)."u";}

					if(substr($value,-1,1) eq "n" and substr($Vlist[2],-2,1) eq "n") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2))."n";}
					if(substr($value,-1,1) eq "n" and substr($Vlist[2],-2,1) eq "m") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000000)."n";}
					if(substr($value,-1,1) eq "n" and substr($Vlist[2],-2,1) eq "u") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000)."n";}
					if(substr($value,-1,1) eq "n" and substr($Vlist[2],-2,1) eq "p") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)/1000)."n";}

					if(substr($value,-1,1) eq "p" and substr($Vlist[2],-2,1) eq "p") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2))."p";}
					if(substr($value,-1,1) eq "p" and substr($Vlist[2],-2,1) eq "u") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000000)."p";}
					if(substr($value,-1,1) eq "p" and substr($Vlist[2],-2,1) eq "n") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000)."p";}

					if(substr($value,-1,1) =~ '\d' and substr($Vlist[2],-2,1) =~ '\d') {$value = ($value * substr($Vlist[2],0,-1))/($value + substr($Vlist[2],0,-1));}		# ok
					if(substr($value,-1,1) =~ '\d' and substr($Vlist[2],-2,1) eq "k") {$value = ($value * substr($Vlist[2],0,-2)*1000)/($value + substr($Vlist[2],0,-2)*1000);}	#ok
					if(substr($value,-1,1) eq "k" and substr($Vlist[2],-2,1) =~ '\d') {$value = ((substr($value,0,-1)*1000) * substr($Vlist[2],0,-1))/((substr($value,0,-1)*1000) + substr($Vlist[2],0,-1));}	#ok
					
					if(substr($value,-1,1) eq "k" and substr($Vlist[2],-2,1) eq "k") {$value = (substr($value,0,-1) * substr($Vlist[2],0,-2))/(substr($value,0,-1) + substr($Vlist[2],0,-2))."k";}	#ok
					if(substr($value,-1,1) eq "k" and substr($Vlist[2],-2,1) eq "M") {$value = (substr($value,0,-1) * substr($Vlist[2],0,-2)*1000)/(substr($value,0,-1) + substr($Vlist[2],0,-2)*1000)."k";}	#ok
					if(substr($value,-1,1) eq "M" and substr($Vlist[2],-2,1) eq "k") {$value = ((substr($value,0,-1)*1000) * substr($Vlist[2],0,-2))/((substr($value,0,-1)*1000) + substr($Vlist[2],0,-2))."k";}

					if(substr($value,-1,1) eq "M" and substr($Vlist[2],-2,1) eq "M") {$value = (substr($value,0,-1) * substr($Vlist[2],0,-2))/(substr($value,0,-1) + substr($Vlist[2],0,-2))."M";}
					}
				if(substr($value,0,-1) >= 1000 and substr($value,-1,1) eq "p"){$value = sprintf("%.3f",(substr($value,0,-1)/1000))."n";}
				if(substr($value,0,-1) >= 1000 and substr($value,-1,1) eq "n"){$value = sprintf("%.3f",(substr($value,0,-1)/1000))."u";}
				if(substr($value,0,-1) >= 1000 and substr($value,-1,1) eq "u"){$value = sprintf("%.3f",(substr($value,0,-1)/1000))."m";}

				if(substr($value,-1,1) =~ '\d' and $value < 1000){$value = sprintf("%.3f",$value);}			#ok
				if(substr($value,-1,1) eq "k" and $value =~ /\./){$value = sprintf("%.3f",substr($value,0,-1)); $value =~ s/\.?0+$//; $value = $value."k";}	#ok
				if(substr($value,-1,1) eq "M" and $value =~ /\./){$value = sprintf("%.3f",substr($value,0,-1)); $value =~ s/\.?0+$//; $value = $value."M";}	#ok

				while($value){if($value =~ '\.' and substr($value,-1,1) =~ '\d' and substr($value,-1,1) eq '0' ) {$value =  substr($value,0,-1); } else{last;}}
				while($value){if($value =~ '\.' and substr($value,-1,1) =~ '\D' and substr($value,-2,1) eq '0' ) {$value =  substr($value,0,-2).substr($value,-1,1); } elsif(substr($value,-2,1) =~ '\.' and substr($value,-1,1) =~ '\D' ) {$value =  substr($value,0,-2).substr($value,-1,1); } else{last;}}
				}
			
			if ($list[0] =~ /^(resistor|capacitor|inductor|zener)$/)						## matching type ##
			{
				$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
				$tested-> write($rowT, 1, $list[0], $format_data);  						## TestType ##
				$tested-> write($rowT, 2, $value, $format_data);  							## BOMvalue ##

				$tested->conditional_formatting($rowT, 3,
				{
					type     => 'cell',
					criteria => 'not equal to',
					value    => $value,
					format   => $format_VCC,
					});

				@param =  split('\,', $list[1]);
				while($param[0]){if($param[0] =~ '\.' and substr($param[0],-1,1) =~ '\d' and substr($param[0],-1,1) eq '0' ) {$param[0] =  substr($param[0],0,-1); } else{last;}}
				while($param[0]){if($param[0] =~ '\.' and substr($param[0],-1,1) =~ '\D' and substr($param[0],-2,1) eq '0' ) {$param[0] =  substr($param[0],0,-2).substr($param[0],-1,1); } elsif(substr($param[0],-2,1) =~ '\.' and substr($param[0],-1,1) =~ '\D' ) {$param[0] =  substr($param[0],0,-2).substr($param[0],-1,1); } else{last;}}
				$tested-> write($rowT, 3, $param[0], $format_data);  						## Nominal ## 
				if($param[1] <= 40){$tested-> write($rowT, 4, $param[1], $format_data);}  	## HiLimit ##
				if($param[1] > 40){$tested-> write($rowT, 4, $param[1], $format_STP);}  	## HiLimit ##
				if($param[2] <= 40){$tested-> write($rowT, 5, $param[2], $format_data);}  	## LoLimit ##
				if($param[2] > 40){$tested-> write($rowT, 5, $param[2], $format_STP);}		## LoLimit ##
				$tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);  	## Comment ##
				$rowT++;
				goto Next_Ori;
				}
			elsif ($list[0] eq "diode")														## matching type ##
			{
				$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
				$tested-> write($rowT, 1, $list[0], $format_data);  						## TestType ##
				@param =  split('\,', $list[1]);
				if ($lineTF !~ m/\"/g)
				{
					$tested-> write($rowT, 4, $param[0], $format_data);  					## HiLimit ##
					$tested-> write($rowT, 5, $param[1], $format_data);	  					## LoLimit ##
					$tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);  ## Comment ##
					}
				if ($lineTF =~ m/\"/g)
				{
					$DioNom = "";	$DioNom =  $DioNom . $param[0];  						## Nominal ## 
					$DioHiL = "";	$DioHiL =  $DioHiL . $param[1];  						## HiLimit ##
					$DioLoL = "";	$DioLoL =  $DioLoL . $param[2];	  						## LoLimit ##
					$tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);  ## Comment ##
					while($lineTF = <SourceFile>)
					{
						chomp($lineTF);
						$lineTF =~ s/^ +//;
						if (substr($lineTF,0,5) eq "diode")
						{
							#print $lineTF."\n";
							@param =  split('\,', substr($lineTF,6));
							$DioNom =  $DioNom . "\n" . $param[0];  						## Nominal ## 
							$DioHiL =  $DioHiL . "\n" . $param[1];  						## HiLimit ##
							$DioLoL =  $DioLoL . "\n" . $param[2];	  						## LoLimit ##
							}
						elsif (eof)
						{
							$tested-> write($rowT, 3, $DioNom, $format_data);  				## Nominal ## 
							$tested-> write($rowT, 4, $DioHiL, $format_data);  				## HiLimit ##
							$tested-> write($rowT, 5, $DioLoL, $format_data);	  			## LoLimit ##
							last;
							}
						}
					}
				$rowT++;
				goto Next_Ori;
				}
			elsif ($list[0] eq "jumper")													## matching type ##
			{
				$state = shorts_check($ibus, $sbus);
				$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
				$tested-> write($rowT, 1, $list[0], $format_data);  						## TestType ##
				if($value eq '-'){$tested-> write($rowT, 2, $value, $format_data);}  		## BOMvalue ##
				elsif((substr($value,-1,1) =~ '\d') and $value < 10){$tested-> write($rowT, 2, $value, $format_data);}
				else{$tested-> write($rowT, 2, $value, $format_VCC);}
				@param =  split('\,', $list[1]);
				$OP = 0; if ($lineTF =~ "op") {$OP = 1;}

				if ($OP == 0 and $state eq "True"){$tested-> write($rowT, 4, $param[0], $format_data); $tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);}
				if ($OP == 0 and $state eq "False"){$tested-> write($rowT, 4, $param[0], $format_STP); $tested-> write($rowT, 6, "pair bus not cover in 'Shorts'", $format_VCC);}
				if ($OP == 1){$tested-> write($rowT, 5, $param[0], $format_data);}
				if ($OP == 1){$tested-> write($rowT, 6, "OP test", $format_STP);}
				$rowT++;
				goto Next_Ori;
				}
			elsif ($list[0] eq "fuse")					  									## matching type ##
			{
				$tested-> write($rowT, 0, $device, $format_item);  							## TestName ##
				$tested-> write($rowT, 1, $list[0], $format_data);  						## TestType ##
				@param =  split('\,', $list[1]);
				$tested-> write($rowT, 4, $param[0], $format_data);
				$tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);  	## Excel ##
				$rowT++;
				goto Next_Ori;
				}
			elsif (eof and $foundTO == 1)					  								## no parameter ##
			{
				$tested-> write($rowT, 0, $device, $format_item); 									## Excel ##
				$tested-> write($rowT, 1, "[base ver]", $format_data);  							## TestType ##
				$tested-> write($rowT, 6, "No Test Parameter Found in TestFile.", $format_anno1);  	## Excel ##
				$rowT++;
				goto Next_Ori;
				}
			}}
		Next_Ori:
	################ testable device [sub-versions] ######################################
	for ($v = 0; $v < $len_ver; $v = $v + 1)
	{
		if (exists($testorder{$versions[$v]."+".$device}))
		{
			if(substr($testplan{$device},0,6) eq "tested"
			and grep{ $_ eq $testorder{$versions[$v]."+".$device}} @TOtestedver)
			{
				$foundTO = 1;
				print "			Common AnaVTest		", $device." - [$versions[$v]]\n";
				@array = ("-","-","-","-","-");
				$array_ref = \@array;
				$tested-> write_row($rowT, 2, $array_ref, $format_data);
			
				$UNCover = 1;
				$coverage-> write($rowC, 1, 'V', $format_data);			#Coverage
	
				open(SourceFile, "<$versions[$v]/analog/$device") or open(SourceFile, "<$versions[$v]/analog/1%$device") or warn "\t!!! Failed to open '$device' file: $!.\n";
				if ($! eq "No such file or directory"){$tested-> write($rowT, 6, "TestFile not found.", $format_anno1);}
				else
				{
				while($lineTF = <SourceFile>)							#read parameter
				{
					$len = 0;
					chomp($lineTF);
					$lineTF =~ s/^ +//;                               	#clear head of line spacing
					#print $lineTF,"\n";
					next if (length($lineTF) < 2);
					my @list = split('\ ', $lineTF, 2);

					if ($list[0] eq "connect" and substr($list[1],0,1) eq "i" and $list[1] !~ " pins "){my @list1 = split('\"', $list[1], 3); $ibus = uc($list1[1]); }
					if ($list[0] eq "connect" and substr($list[1],0,1) eq "s" and $list[1] !~ " pins "){my @list1 = split('\"', $list[1], 3); $sbus = uc($list1[1]); }

					if ($list[0] =~ /^(resistor|capacitor|inductor|zener)$/)							## matching type ##
					{
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ".$list[0], $format_data);  	## TestType ##
						@param =  split('\,', $list[1]);
						$tested-> write($rowT, 3, $param[0], $format_data);  							## Nominal ## 
						if($param[1] <= 40){$tested-> write($rowT, 4, $param[1], $format_data);}  		## HiLimit ##
						if($param[1] > 40){$tested-> write($rowT, 4, $param[1], $format_STP);}  		## HiLimit ##
						if($param[2] <= 40){$tested-> write($rowT, 5, $param[2], $format_data);}  		## LoLimit ##
						if($param[2] > 40){$tested-> write($rowT, 5, $param[2], $format_STP);}			## LoLimit ##
						$tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);  		## Comment ##
						$rowT++;
						goto Next_Rev;
						}
					elsif ($list[0] eq "diode")															## matching type ##
					{
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ".$list[0], $format_data);  	## TestType ##
						@param =  split('\,', $list[1]);
						if ($lineTF !~ m/\"/g){
						$tested-> write($rowT, 4, $param[0], $format_data);  							## HiLimit ##
						$tested-> write($rowT, 5, $param[1], $format_data);	  							## LoLimit ##
						$tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);  		## Comment ##
						}
						if ($lineTF =~ m/\"/g)
						{
							$DioNom = "";	$DioNom =  $DioNom . $param[0];  							## Nominal ## 
							$DioHiL = "";	$DioHiL =  $DioHiL . $param[1];  							## HiLimit ##
							$DioLoL = "";	$DioLoL =  $DioLoL . $param[2];	  							## LoLimit ##
							$tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);  	## Comment ##
							while($lineTF = <SourceFile>)
							{
								chomp($lineTF);
								$lineTF =~ s/^ +//;
								if (substr($lineTF,0,5) eq "diode")
								{
									@param =  split('\,', substr($lineTF,6));
									$DioNom =  $DioNom . "\n" . $param[0];  							## Nominal ## 
									$DioHiL =  $DioHiL . "\n" . $param[1];  							## HiLimit ##
									$DioLoL =  $DioLoL . "\n" . $param[2];	  							## LoLimit ##
									}
								elsif (eof)
								{
									$tested-> write($rowT, 3, $DioNom, $format_data);  					## Nominal ## 
									$tested-> write($rowT, 4, $DioHiL, $format_data);  					## HiLimit ##
									$tested-> write($rowT, 5, $DioLoL, $format_data);	  				## LoLimit ##
									last;
									}
								}
							}
						$rowT++;
						goto Next_Rev;
						}
					elsif ($list[0] eq "jumper")														## matching type ##
					{
						$state = shorts_check($ibus, $sbus);
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ".$list[0], $format_data);  	## TestType ##
						@param =  split('\,', $list[1]);
						$OP = 0; if ($lineTF =~ "op") {$OP = 1;}

						if ($OP == 0 and $state eq "True"){$tested-> write($rowT, 4, $param[0], $format_data); $tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);}
						if ($OP == 0 and $state eq "False"){$tested-> write($rowT, 4, $param[0], $format_STP); $tested-> write($rowT, 6, "pair bus not cover in 'Shorts'", $format_VCC);}
						if ($OP == 1){$tested-> write($rowT, 5, $param[0], $format_data);}
						if ($OP == 1){$tested-> write($rowT, 6, "OP test", $format_STP);}
						$rowT++;
						goto Next_Rev;
						}
					elsif ($list[0] eq "fuse")					  										## matching type ##
					{
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ".$list[0], $format_data);  	## TestType ##
						@param =  split('\,', $list[1]);
						$tested-> write($rowT, 4, $param[0], $format_data);
						$tested-> write($rowT, 6, substr($testplan{$device},6), $format_anno);  		## Excel ##
						$rowT++;
						goto Next_Rev;
						}
					elsif (eof and $foundTO == 1)					  										## no parameter ##
					{
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ", $format_STP);  					## TestType ##
						$tested-> write($rowT, 6, "No Test Parameter Found in TestFile.", $format_anno1);  	## Excel ##
						$rowT++;
						goto Next_Rev;
						}
					}}
				Next_Rev:
				}
	################ testorder skipped devices [sub-version] #############################
			elsif(grep{ $_ eq $testorder{$versions[$v]."+".$device}} @TOuntestver)
			{
				$foundTO = 1;
				print "			Common NullTest		", $device." - [$versions[$v]]\n";
				$tested-> write($rowT, 0, $device, $format_data);
				$tested-> write($rowT, 1, "[".$versions[$v]."] - ", $format_STP);  							## TestType ##
					@array = ("-","-","-","-");
					$array_ref = \@array;
					$tested-> write_row($rowT, 2, $array_ref, $format_data);
				$tested-> write($rowT, 6, "test is skipped in version [$versions[$v]]", $format_anno1);  	## Excel ##
				$rowT++;
				}
			}
		}
	if ($rowT - $row_ver > 1){$tested-> merge_range($row_ver, 0, $rowT-1, 0, $device, $format_item);}
	}

	################ testplan skipped devices ############################################
	elsif (exists($testplan{$device})
	and substr($testplan{$device},0,7) eq "skipped"
	and grep{ $_ eq $testorder{$device}} @TOtested)
	{
		$foundTO = 1;
		print "			Common SkipTest		", $device,"\n";
		if ($UNCover == 0){$coverage-> write($rowC, 1, 'K', $format_VCC);}					#Coverage
		$untest-> write($rowU, 0, $device, $format_data);
		$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_anno1);  			## Excel ##
		$untest-> write($rowU, 2, substr($testplan{$device},8), $format_anno);  			## Excel ##
		$rowU++;
		next; #goto Next_Dev;
		}
	################ testplan inexistent devices #########################################
	elsif (not exists($testplan{$device})
	and exists($testorder{$device})
	and grep{ $_ eq $testorder{$device}} @TOtested)
	{
		$foundTO = 1;
		print "			Testplan InExistent	", $device,"\n";
		if ($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_data);}					#Coverage
		$untest-> write($rowU, 0, $device, $format_data);									## Excel ##
		$untest-> write($rowU, 1, "Unidentified in Testplan.", $format_VCC);
		$untest-> write($rowU, 2, $testorder{$device}, $format_anno);						## Excel ##
		$rowU++;
		next; #goto Next_Dev;
		}
	################ testorder Nulltest devices ##########################################
	elsif(exists($testorder{$device})
	and grep{ $_ eq $testorder{$device}} @TOuntest)
	{
		$foundTO = 1;
		print "			Common NullTest		", $device,"\n";
		$untest-> write($rowU, 0, $device, $format_data);
		$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_anno);
		$UTline = "";
		$fileF = 0;
		
		if($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_data);}					#Coverage
		open(ALL, "<analog/$device") or open(ALL, "<analog/1%$device") or warn "\t!!! Failed to open '$device' file: $!.\n";
		if ($! eq "No such file or directory"){$untest-> write($rowU, 2, "TestFile not found.", $format_anno1);}
		else
		{
		while($line = <ALL>)
		{
			$fileF = 1;
			if (index($line,$device)>1)
			{
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;
				}
			elsif (eof){last;}
			}}
		$UTline =~ s/(^\s+|\s+$)//g;
		if($UTline eq "" and $fileF == 1){$untest-> write($rowU, 2, "No Comments Found in TestFile.", $format_anno1);}
		if($UTline ne ""){
		$untest-> write($rowU, 2, $UTline, $format_anno);
		$untest-> set_column(2, 2, $length_anno);}
		$rowU++;
		close ALL;
		#next; #goto Next_Dev;
		}
	################ testorder skipped devices ###########################################
	elsif(exists($testorder{$device})
	and grep{ $_ eq $testorder{$device}} @TOskipped)
	{
		$foundTO = 1;
		print "			Common Skipped		", $device,"\n";
		$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
		$untest-> write($rowU, 1, "been Skipped in TestOrder.", $format_anno1);  			## Excel ##
		$UTline = "";
		$fileF = 0;
		
		if($UNCover == 0){$coverage-> write($rowC, 1, 'K', $format_VCC);}					#Coverage
		open(ALL, "<analog/$device") or open(ALL, "<analog/1%$device") or warn "\t!!! Failed to open '$device' file: $!.\n";
		if ($! eq "No such file or directory"){$untest-> write($rowU, 2, "TestFile not found.", $format_anno1);}
		else
		{
		while($line = <ALL>)
		{
			$fileF = 1;
			if (index($line,$device)>1)
			{
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;
				}
			elsif (eof){last;}
			}}
		$UTline =~ s/(^\s+|\s+$)//g;
		if($UTline eq "" and $fileF == 1){$untest-> write($rowU, 2, "No Comments Found in TestFile.", $format_anno1);}
		if($UTline ne ""){
		$untest-> write($rowU, 2, $UTline, $format_anno);
		$untest-> set_column(2, 2, $length_anno);}
		$rowU++;
		close ALL;
		#next; #goto Next_Dev;
		}
	################ parallel tested devices #############################################
	elsif(exists($testorder{$device})
	and grep{ $_ eq $testorder{$device}} @TOparal)
	{
		$foundTO = 1;
		print "			Common ParalTest	", $device,"\n";
		
		if($UNCover == 0){$coverage-> write($rowC, 1, 'L', $format_data);}			#Coverage
		$limited-> write($rowL, 0, $device, $format_data);		## Excel ##
		#$limited-> write($rowL, 1, $dev[2], $format_anno);		## Excel ##
		$UTline = "";
		$fileF = 0;
		
		open(ALL, "<analog/$device") or open(ALL, "<analog/1%$device") or warn "\t!!! Failed to open '$device' file: $!.\n";
		if ($! eq "No such file or directory"){$limited-> write($rowL, 1, "TestFile not found.", $format_anno1);}
		else
		{
		while($line = <ALL>)
		{
			$fileF = 1;
			if (index($line,$device)>1)
			{
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;
				}
			elsif (eof){last;}
			}}
		$UTline =~ s/(^\s+|\s+$)//g;
		if($UTline eq "" and $fileF == 1){$limited-> write($rowL, 1, "No Comments Found in TestFile.", $format_anno1);}
		if($UTline ne ""){
		$limited-> write($rowL, 1, $UTline, $format_anno);	
		$limited-> set_column(1, 1, $length_anno);}
		$rowL++;
		close ALL;
		goto Next_Dev;
		}
	################ testable analog powered test ########################################
	elsif(exists($testorder{$device})
	and ($testorder{$device} eq "tested-pwr" or $testorder{$device} eq "untest-pwr")
	)
	{
		$foundTO = 1;
		print "			Common PwrTest		", $device,"\n";
		@array = ("-","-","-","-","-","-","-","-","-","-");
		$array_ref = \@array;
		$power-> write_row($rowP, 3, $array_ref, $format_data);

		if($testorder{$device} eq "tested-pwr")
		{
			if (exists($testplan{$device}) and substr($testplan{$device},0,6) eq "tested")
			{
				$cover = 1;
				$coverage-> write($rowC, 3, 'V', $format_data);											# Coverage
				$power-> write($rowP, 0, $device, $format_data);										## Excel ##
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$device, $format_anno);							## Excel ##
				}
			elsif (exists($testplan{$device}) and substr($testplan{$device},0,7) eq "skipped")
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'K', $format_VCC);}						#Coverage
				$power-> write($rowP, 0, $device, $format_data);										## Excel ##
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);		## Excel ##
				}
			elsif(not exists($testplan{$device}))
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}						#Coverage
				$power-> write($rowP, 0, $device, $format_data);										## Excel ##
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Unidentified in Testplan.", $format_VCC);
				}
			
		if (exists($testplan{$device}) and substr($testplan{$device},0,6) eq "tested")
		{
			open(ALL, "<analog/$device") or open(ALL, "<analog/1%$device") or  warn "\t!!! Failed to open '$device' file: $!.\n";
			if ($! eq "No such file or directory"){$power-> write($rowP, 3, "TestFile not found.", $format_anno1);}
			else
			{
				@testname = ();
				while($line = <ALL>)
				{
					chomp ($line);
					$line =~ s/(^\s+|\s+$)//g;
					next if ($line eq "");
					#print $line,"\n";
					my @list = split('\"', $line);
					$list[0] =~ s/(^\s+|\s+$)//g;
					if ($list[0] eq "test")
					{
						#print $list[1],"\n";
						push(@testname, uc($list[1]));
						}
					last if ($line =~ "end test")
					}
				while($line = <ALL>)
				{
					chomp ($line);
					$line =~ s/(^\s+|\s+$)//g;
					next if ($line eq "");
					my @list = split('\"', $line);
					$list[0] =~ s/(^\s+|\s+$)//g;
					if ($list[0] eq "subtest")
					{
						foreach my $i (0..@testname-1)
						{
							if (grep{ $_ eq uc($list[1])} @testname and uc($list[1]) eq uc($testname[$i]))
							{
								while($line = <ALL>)
								{
									$line =~ s/(^\s+|\s+$)//g;
									if (substr($line,0,7) eq "measure")
									{
										#print $testname[$i],"\n";
										$testname[$i]= $testname[$i]." / ".$line."\n";
										#print $testname[$i],"\n";
										goto OUTER;
										}
									}
								}
							}
						}
					OUTER:
					}
				}
			close ALL;
			$content = join("",@testname);
			$content =~ s/(^\s+|\s+$)//g;
			$power-> write($rowP, 3, $content, $format_anno);
			if($content eq ""){$power-> write($rowP, 3, "NO valid test found.", $format_anno1);}
			}
		$rowP++;
		}
			
		elsif($testorder{$device} eq "untest-pwr")
		{
			if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}				#Coverage
			$untest-> write($rowU, 0, $device, $format_data);								## Excel ##
			$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  		## Excel ##
			$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  					## Excel ##
			$rowU++;
			}
		}
	################ analog powered [sub-version] ########################################
	for ($v = 0; $v < $len_ver; $v = $v + 1)
	{
		if (exists($testorder{$versions[$v]."+".$device}) and substr($testorder{$versions[$v]."+".$device},7,7) eq "pwr+ver")
		{
		$foundTO = 1;
		print "			Common PwrVTest		", $device." - [$versions[$v]]\n";
		@array = ("-","-","-","-","-","-","-","-","-","-");
		$array_ref = \@array;
		$power-> write_row($rowP, 3, $array_ref, $format_data);

			if ($testorder{$versions[$v]."+".$device} eq "tested-pwr+ver")
			{

				if (exists($testplan{$device}) and substr($testplan{$device},0,6) eq "tested")
				{
					$cover = 1;
					$coverage-> write($rowC, 3, 'V', $format_data);												#Coverage
					$power-> write($rowP, 1, "tested-pwr - ".$device." - [".$versions[$v]."]", $format_anno);	## Excel ##
					$power-> write($rowP, 2, "Tested - ".$device, $format_anno);								## Excel ##
					}
				elsif(exists($testplan{$device}) and substr($testplan{$device},0,7) eq "skipped")
				{
					if ($cover == 0){$coverage-> write($rowC, 3, 'K', $format_VCC);}							#Coverage
					$power-> write($rowP, 1, "tested-pwr - ".$device." - [".$versions[$v]."]", $format_anno);	## Excel ##
					$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);			## Excel ##
					}
				elsif(not exists($testplan{$device}))
				{
					if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}							#Coverage
					$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);					## Excel ##
					$power-> write($rowP, 2, "Unidentified in Testplan.", $format_VCC);
					}
				
			if (exists($testplan{$device}) and substr($testplan{$device},0,6) eq "tested")
			{
				open(ALL, "<analog/$device") or open(ALL, "<analog/1%$device") or  warn "\t!!! Failed to open '$device' file: $!.\n";
				if ($! eq "No such file or directory"){$power-> write($rowP, 3, "TestFile not found.", $format_anno1);}
				else
				{
					@testname = ();
					while($line = <ALL>)
					{
						chomp ($line);
						$line =~ s/(^\s+|\s+$)//g;
						next if ($line eq "");
						#print $line,"\n";
						my @list = split('\"', $line);
						$list[0] =~ s/(^\s+|\s+$)//g;
						if ($list[0] eq "test")
						{
							#print $list[1],"\n";
							push(@testname, uc($list[1]));
							}
						last if ($line =~ "end test")
						}
					while($line = <ALL>)
					{
						chomp ($line);
						$line =~ s/(^\s+|\s+$)//g;
						next if ($line eq "");
						my @list = split('\"', $line);
						$list[0] =~ s/(^\s+|\s+$)//g;
						if ($list[0] eq "subtest")
						{
							foreach my $i (0..@testname-1)
							{
								if (grep{ $_ eq uc($list[1])} @testname and uc($list[1]) eq uc($testname[$i]))
								{
									while($line = <ALL>)
									{
										$line =~ s/(^\s+|\s+$)//g;
										if (substr($line,0,7) eq "measure")
										{
											#print $testname[$i],"\n";
											$testname[$i]= $testname[$i]." / ".$line."\n";
											#print $testname[$i],"\n";
											goto OUTER;
											}
										}
									}
								}
							}
						OUTER:
						}
					}
				close ALL;
				$content = join("",@testname);
				$content =~ s/(^\s+|\s+$)//g;
				$power-> write($rowP, 3, $content, $format_anno);
				if($content eq ""){$power-> write($rowP, 3, "NO valid test found.", $format_anno1);}
				}
			$rowP++;
			}

			elsif($testorder{$versions[$v]."+".$device} eq "untest-pwr+ver")
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}						#Coverage
				$power-> write($rowP, 1, "Skipped - $device -ver: [".$versions[$v]."]", $format_anno1);	## Excel ##
				$power-> write($rowP, 2, "-", $format_data);											## Excel ##
				$rowP++;
				}
			}
		}
	################ testable digital test ###############################################
	if(exists($testorder{$device})
	and ($testorder{$device} eq "tested-dig" or $testorder{$device} eq "untest-dig")
	)
	{			
		$foundTO = 1;
		$length_DigPin = 10;
		print "			Common DigiTest		", $device,"\n";
		@array = ("-","-","-","-","-","-","-","-","-","-");
		$array_ref = \@array;
		$power-> write_row($rowP, 3, $array_ref, $format_data);
	
		if($testorder{$device} eq "tested-dig")
		{
			if (exists($testplan{$device}) and substr($testplan{$device},0,6) eq "tested")
			{
				$cover = 1;
				$coverage-> write($rowC, 2, 'V', $format_data);										#Coverage
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);			## Excel ##
				$power-> write($rowP, 2, "Tested - ".$device, $format_anno);						## Excel ##
				}
			elsif (exists($testplan{$device}) and substr($testplan{$device},0,7) eq "skipped")
			{
				if ($cover == 0){$coverage-> write($rowC, 2, 'K', $format_VCC);}					#Coverage
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);			## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);	## Excel ##
				}
			elsif(not exists($testplan{$device}))
			{
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_data);}					#Coverage
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);			## Excel ##
				$power-> write($rowP, 2, "Unidentified in Testplan.", $format_VCC);
				}
			
			###### hyperlink #####################################
			$power-> write_url($rowP, 0, 'internal:'.$device.'!A1');	## hyperlink
			
			if ($worksheet == 0){
			$worksheet = 1;
			$IC = $bom_coverage_report-> add_worksheet($device);		## hyperlink
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
						#print $lineDig."\n";
						if ($lineDig eq uc($device))
						{
							while($lineDig = <Boards>) 
							{
								#print $lineDig;
								$Total_Pin++;
								@DigPin = split('\.',$lineDig);
								$DigPin[0] =~ s/(^\s+|\s+$)//g;
								$DigPin[1] =~ s/(^\s+|\s+$)//g;
								#print $DigPin[0]."\n";
								if ($DigPin[1] =~ /(GND|GROUND)/)
								{
									if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_GND); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
									if ($DigPin[0] =~ /^\D/i)
									{
										$IC-> write($DigPin[0], $DigPin[1], $format_GND);
										($pos) = $DigPin[0] =~ /^\D+/g;
										if (length($pos) == 1){$DigPos = ord($pos)%64;}
										if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
										#print $DigPos."\n";
										if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
										}
									$GND_Pin++;
									}
								#elsif ($DigPin[1] =~ /(^\+0|^0V|^\+1|^1V|^\+2|^2V|^\+3|^3V|^\+5|^5V|^V_|^VCC|^VDD|^PP|^P0V|^P1V|^P2V|^P3V|^P5V)/)
								elsif ($DigPin[1] =~ /^(\+0|0V|\+1|1V|\+2|2V|\+3|3V|\+5|5V|V_|VCC|VDD|PP|P0V|P1V|P2V|P3V|P5V)/)
								{
									if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_VCC); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
									if ($DigPin[0] =~ /^\D/i)
									{
										$IC-> write($DigPin[0], $DigPin[1], $format_VCC);
										($pos) = $DigPin[0] =~ /^\D+/g;
										if (length($pos) == 1){$DigPos = ord($pos)%64;}
										if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
										#print $DigPos."\n";
										if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
										}
									$Power_Pin++;
									}
								elsif ($DigPin[1] =~ /(^NC_|_NC$|NONE)/)
								{
									if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_NC); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
									if ($DigPin[0] =~ /^\D/i)
									{
										$IC-> write($DigPin[0], $DigPin[1], $format_NC);
										($pos) = $DigPin[0] =~ /^\D+/g;
										if (length($pos) == 1){$DigPos = ord($pos)%64;}
										if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
										#print $DigPos."\n";
										if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
										}
									$NC_Pin++;
									}
								else{
									if (exists($bdg_list{uc($device)."\.".$DigPin[0]}))
									{
										#print uc($device)."\.".$DigPin[0],"\n";
										if($bdg_list{uc($device)."\.".$DigPin[0]})
										{
											#print uc($device)."\.".$DigPin[0],"\n";
											if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Toggle_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
											if ($DigPin[0] =~ /^\D/i)
											{
												$IC-> write($DigPin[0], $DigPin[1]."\n* Toggle_Test", $format_data);
												($pos) = $DigPin[0] =~ /^\D+/g;
												if (length($pos) == 1){$DigPos = ord($pos)%64;}
												if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
												if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
												}
											}
										}
									else{
										if(exists($hash_pin{$DigPin[1]}))
										{
											if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Contact_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
											if ($DigPin[0] =~ /^\D/i)
											{
												$IC-> write($DigPin[0], $DigPin[1]."\n* Contact_Test", $format_data);
												($pos) = $DigPin[0] =~ /^\D+/g;
												if (length($pos) == 1){$DigPos = ord($pos)%64;}
												if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
												#print $DigPin[1]."\n";
												if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
												}
											}
										else{
											if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1],$format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
											if ($DigPin[0] =~ /^\D/i)
											{
												$IC-> write($DigPin[0], $DigPin[1], $format_data);
												($pos) = $DigPin[0] =~ /^\D+/g;
												if (length($pos) == 1){$DigPos = ord($pos)%64;}
												if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
												if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
												}
											}
										}
									}
								if ($lineDig =~ "\;")
								{
									$power-> write($rowP, 4, $Total_Pin, $format_item);
									$power-> write($rowP, 5, $Power_Pin, $format_VCC);
									$power-> write($rowP, 6, $GND_Pin, $format_GND);
									$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "*Toggle_Test*")', $format_data);
									$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "*Contact_Test*")', $format_data);
									$power-> write($rowP, 9, $NC_Pin, $format_NC);
									$power-> write_formula($rowP, 10, "=(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-H".($rowP+1)."-I".($rowP+1)."-J".($rowP+1).")", $format_data);
									$power-> write_formula($rowP, 11, "=(H".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
									$power-> write_formula($rowP, 12, "=(I".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
									last;
									}
								}
							}
						}
					}
  				}
  			close Boards;}

			###### hyperlink #####################################
			$power-> write($rowP, 0, $device, $format_hylk);

			$family = "";
  			open(SourceFile, "<digital/$device") or open(SourceFile, "<digital/1%$device") or warn "\t!!! Failed to open '$device' file: $!.\n";
			if ($! eq "No such file or directory"){$power-> write($rowP, 3, "TestFile not found.", $format_anno1);}
			else
			{
				while($lineTF = <SourceFile>)							#reading family
				{
					#chomp($lineTF);
					$lineTF =~ s/(^\s+)//g;								#clear head of line spacing
					#print $lineTF;
					if (substr($lineTF,0,6) eq "family")
					{$family = $lineTF . $family;}
  					}
  				close SourceFile;
  				chomp($family);
  				$power-> write($rowP, 3, $family, $format_anno);
  				if($family eq ""){$power-> write($rowP, 3, "Family not define.", $format_anno1);}
  				}
			$rowP++;
			}

		elsif($testorder{$device} eq "untest-dig")
		{
			if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_data);}				#Coverage
			$untest-> write($rowU, 0, $device, $format_data);								## Excel ##
			$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  		## Excel ##
			$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  					## Excel ##
			$rowU++;
			}
		}
	################ digital [sub-version] ###############################################
	for ($v = 0; $v < $len_ver; $v = $v + 1)
	{
		if (exists($testorder{$versions[$v]."+".$device}) and substr($testorder{$versions[$v]."+".$device},7,7) eq "dig+ver")
		{
		$foundTO = 1;
		print "			Common DigVTest		",$device," - [$versions[$v]]\n";
		@array = ("-","-","-","-","-","-","-","-","-","-");
		$array_ref = \@array;
		$power-> write_row($rowP, 3, $array_ref, $format_data);

			if ($testorder{$versions[$v]."+".$device} eq "tested-dig+ver")
			{
				if(exists($testplan{$device}) and substr($testplan{$device},0,6) eq "tested")
				{
					$cover = 1;
					$coverage-> write($rowC, 2, 'V', $format_data);												#Coverage
					$power-> write($rowP, 1, "tested-dig - ".$device." - [".$versions[$v]."]", $format_anno);	## Excel ##
					$power-> write($rowP, 2, "Tested - ".$device, $format_anno);								## Excel ##
					}
				elsif(exists($testplan{$device}) and substr($testplan{$device},0,7) eq "skipped")
				{
					if ($cover == 0){$coverage-> write($rowC, 2, 'K', $format_VCC);}					#Coverage
					$power-> write($rowP, 1, "tested-dig - ".$device." - [".$versions[$v]."]", $format_anno);	## Excel ##
					$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);	## Excel ##
					}
				elsif(not exists($testplan{$device}))
				{
					if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_data);}					#Coverage
					$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);			## Excel ##
					$power-> write($rowP, 2, "Unidentified in Testplan.", $format_VCC);
					}
				$rowP++;
				}
			elsif($testorder{$versions[$v]."+".$device} eq "untest-dig+ver")
			{
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_data);}						#Coverage
				$power-> write($rowP, 1, "Skipped - $device -ver: [".$versions[$v]."]", $format_anno1);	## Excel ##
				$power-> write($rowP, 2, "-", $format_data);											## Excel ##
				$rowP++;
				}
			}
		}
	################ testable mixed test #################################################
	if(exists($testorder{$device})
	and ($testorder{$device} eq "tested-mix" or $testorder{$device} eq "untest-mix")
	)
	{
		$foundTO = 1;
		print "			Common MixTest		", $device,"\n";
		@array = ("-","-","-","-","-","-","-","-","-","-");
		$array_ref = \@array;
		$power-> write_row($rowP, 3, $array_ref, $format_data);

		if($testorder{$device} eq "tested-mix")
		{
			if(exists($testplan{$device}) and substr($testplan{$device},0,6) eq "tested")
			{
				$cover = 1;
				$coverage-> write($rowC, 3, 'V', $format_data);								# Coverage
				$power-> write($rowP, 0, $device, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);	## Excel ##
				$power-> write($rowP, 2, "Tested - ".$device, $format_anno);				## Excel ##
				}
			elsif(exists($testplan{$device}) and substr($testplan{$device},0,7) eq "skipped")
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'K', $format_VCC);}			# Coverage
				$power-> write($rowP, 0, $device, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);		## Excel ##
				}
			elsif(not exists($testplan{$device}))
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}			# Coverage
				$power-> write($rowP, 0, $device, $format_data);							## Excel ##
				$power-> write($rowP, 1, $testorder{$device}." - ".$device, $format_anno);	## Excel ##
				$power-> write($rowP, 2, "Unidentified in Testplan.", $format_VCC);
				}
			$rowP++;
			}
		elsif($testorder{$device} eq "untest-mix")
		{
			if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}				# Coverage
			$untest-> write($rowU, 0, $device, $format_data);								## Excel ##
			$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  		## Excel ##
			$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  					## Excel ##
			$rowU++;
			}
		}
	################ testable Bscan device ###############################################
	elsif(exists($testorder{$device})
	and ($testorder{$device} eq "tested-bscan" or $testorder{$device} eq "untest-bscan")
	)
	{
		$foundTO = 1;
		$length_SNail = 10;
		print "			Common BscTest		", $device,"\n";
		@array = ("-","-","-","-","-","-","-","-","-");
		$array_ref = \@array;
		$power-> write_row($rowP, 4, $array_ref, $format_data);

		if($testorder{$device} eq "tested-dig")
		{

			if(exists($testplan{$device}) and substr($testplan{$device},0,6) eq "tested")
			{
				$cover = 1;
				$coverage-> write($rowC, 4, 'V', $format_data);											#Coverage
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);							## Excel ##
				$power-> write($rowP, 2, "Tested - ".$device, $format_anno);							## Excel ##
				}
			elsif(exists($testplan{$device}) and substr($testplan{$device},0,7) eq "skipped")
			{
				if ($cover == 0){$coverage-> write($rowC, 4, 'K', $format_VCC);}						#Coverage
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);							## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$device},8), $format_anno1);		## Excel ##
				}
			elsif(not exists($testplan{$device}))
			{
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_data);}						#Coverage
				$power-> write($rowP, 1, $testorder{$device}, $format_anno);  							## Excel ##
				$power-> write($rowP, 2, "Unidentified in Testplan.", $format_VCC);
				}

			###### hyperlink #############################################################
			$power-> write_url($rowP, 0, 'internal:'.$device.'!A1');    		## hyperlink

			if ($worksheet == 0)
			{
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
							{
								while($lineDig = <Boards>) 
								{
									$Total_Pin++;
									@BscanNail = split('\.',$lineDig);
									$BscanNail[0] =~ s/(^\s+|\s+$)//g;
									$BscanNail[1] =~ s/(^\s+|\s+$)//g;
									if ($BscanNail[1] =~ /(GND|GROUND)/)
									{
										if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_GND); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
										if ($BscanNail[0] =~ /^\D/i)
										{
											$IC-> write($BscanNail[0], $BscanNail[1], $format_GND);
											($pos) = $BscanNail[0] =~ /^\D+/g;
											if (length($pos) == 1){$NailPos = ord($pos)%64;}
											if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
											if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
											}
										$GND_Pin++;
										}
									elsif ($BscanNail[1] =~ /^(\+0|0V|\+1|1V|\+2|2V|\+3|3V|\+5|5V|V_|VCC|VDD|PP|P0V|P1V|P2V|P3V|P5V)/)
									{
										if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_VCC); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
										if ($BscanNail[0] =~ /^\D/i)
										{
											$IC-> write($BscanNail[0], $BscanNail[1], $format_VCC);
											($pos) = $BscanNail[0] =~ /^\D+/g;
											if (length($pos) == 1){$NailPos = ord($pos)%64;}
											if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
											if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
											}
										$Power_Pin++;
										}
									elsif ($BscanNail[1] =~ /(^NC_|_NC$|NONE)/)
									{
										if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_NC); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
										if ($BscanNail[0] =~ /^\D/i)
										{
											$IC-> write($BscanNail[0], $BscanNail[1], $format_NC);
											($pos) = $BscanNail[0] =~ /^\D+/g;
											if (length($pos) == 1){$NailPos = ord($pos)%64;}
											if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
											if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
											}
										$NC_Pin++;
										}
									else{
										if(exists($hash_pin{$BscanNail[1]})	)
										{
											if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1]."\n* Contact_Test", $format_data); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
											if ($BscanNail[0] =~ /^\D/i)
											{
												$IC-> write($BscanNail[0], $BscanNail[1]."\n* Contact_Test", $format_data);
												($pos) = $BscanNail[0] =~ /^\D+/g;
												if (length($pos) == 1){$NailPos = ord($pos)%64;}
												if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
												if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
												}
											}
										else{
											if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_data); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
											if ($BscanNail[0] =~ /^\D/i)
											{
												$IC-> write($BscanNail[0], $BscanNail[1], $format_data);
												($pos) = $BscanNail[0] =~ /^\D+/g;
												if (length($pos) == 1){$NailPos = ord($pos)%64;}
												if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
												if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
												}
											}
										}
									if ($lineDig =~ "\;")
										{
										$power-> write($rowP, 4, $Total_Pin, $format_item);
										$power-> write($rowP, 5, $Power_Pin, $format_VCC);
										$power-> write($rowP, 6, $GND_Pin, $format_GND);
										$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "*Toggle_Test*")', $format_data);
										$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "*Contact_Test*")', $format_data);
										$power-> write($rowP, 9, $NC_Pin, $format_NC);
										$power-> write_formula($rowP, 10, "=(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-H".($rowP+1)."-I".($rowP+1)."-J".($rowP+1).")", $format_data);
										$power-> write_formula($rowP, 11, "=(H".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
										$power-> write_formula($rowP, 12, "=(I".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
										last;
										}
									}
								}
							}
						}
  					}
  				close Boards;
  				}

			###### hyperlink #############################################################
			$power-> write($rowP, 0, $device, $format_hylk);  				## Excel ##
			
			$family = "";
  			open(SourceFile, "<digital/$device") or open(SourceFile, "<digital/1%$device") or warn "\t!!! Failed to open '$device' file: $!.\n";
			if ($! eq "No such file or directory"){$power-> write($rowP, 3, "TestFile not found.", $format_anno1);}
			else
			{
				while($lineTF = <SourceFile>)							#reading family
				{
					#chomp($lineTF);
					$lineTF =~ s/^ +//;                               	#clear head of line spacing
					#print $lineTF;
					if (substr($lineTF,0,6) eq "family")
					{$family = $lineTF . $family;}
  					#---------------------------------------------------------------------
  					if (substr($lineTF,0,5) eq "nodes")
					{
						while($lineTF = <SourceFile>)
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
  					#---------------------------------------------------------------------
  					}
  				close SourceFile;
  				chomp($family);
  				$power-> write($rowP, 3, $family, $format_anno);
  				if($family eq ""){$power-> write($rowP, 3, "Family not define.", $format_anno1);}
				}
  		$rowP++;
  		}

		elsif($testorder{$device} eq "untest-bscan")
		{
			if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_data);}				#Coverage
			$untest-> write($rowU, 0, $device, $format_data);								## Excel ##
			$untest-> write($rowU, 1, "been skipped in TestOrder.", $format_anno1);  		## Excel ##
			$untest-> write($rowU, 2, $testorder{$device}, $format_anno);  					## Excel ##
			$rowU++;
			}
	
		}

   #~~~~~~~~~~~~~~~ Multiple Test ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	if ($foundTO == 0
	or $testorder{$device} eq "tested-mix" or $testorder{$device} eq "untest-mix"
	or $testorder{$device} eq "tested-pwr" or $testorder{$device} eq "untest-pwr"
	or $testorder{$device} eq "tested-dig" or $testorder{$device} eq "untest-dig"
	or $testorder{$device} eq "tested-bscan" or $testorder{$device} eq "untest-bscan"
	or $testorder{$device} eq "tested-dio" or $testorder{$device} eq "untest-dio")
	{
	#print "--0--",$testorder{$device},"\n";
	my $Mult_file = '';
	for ($i = 0; $i < $sizeTO; $i = $i + 1)
	{
		if (substr($keysTO[$i],0,length($device)) =~ $device and (substr($keysTO[$i],length($device),1) eq "\%" or substr($keysTO[$i],length($device),1) eq "\_"))
		{
			#print "--1--",$keysTO[$i],"\n";
			#print $keysTO[$i],' -- ',$testorder{$keysTO[$i]},"\n";
			#print $testplan{$keysTO[$i]},"\n";
			$Mult_file = $keysTO[$i];
			#print $Mult_file,"\n";
			#print $testorder{$versions[0]."+".$Mult_file},"\n";
	#%%%%%%%%%%%%%%% testable device %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if(exists($testplan{$Mult_file})
	and substr($testplan{$Mult_file},0,6) eq "tested"
	and grep{ $_ eq $testorder{$Mult_file}} @TOtested)
	{
		$foundTO = 1;
		print "			Multiple AnaTest	", $Mult_file."\n";
		@array = ("-","-","-","-","-");
		$array_ref = \@array;
		$tested-> write_row($rowT, 2, $array_ref, $format_data);
		
		$UNCover = 1;	
		$coverage-> write($rowC, 1, 'V', $format_data);			#Coverage
		$row_ver = $rowT;
		
		#$value = read_multi_value(uc($device));
		$value = '-';
		if(exists($bom_value{uc($device)})){$value = $bom_value{uc($device)};} #print "$device value is: $bom_value{uc($device)} \n"; }
		
		open(SourceFile, "<analog/$Mult_file") or open(SourceFile, "<analog/1%$Mult_file") or warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
		if ($! eq "No such file or directory"){$tested-> write($rowT, 6, "TestFile not found.", $format_anno1);}
		else
		{
		while($lineTF = <SourceFile>)							#read parameter
		{
			$len = 0;
			chomp($lineTF);
			$lineTF =~ s/^ +//;                               	#clear head of line spacing
			#print $lineTF,"\n";
			next if (length($lineTF) < 2);
			my @list = split('\ ', $lineTF, 2);
			#print $list[0],"\n";
			#print $list[1],"\n";

			if (grep{ $_ eq $testorder{$Mult_file}} @Paral and $lineTF =~ "report parallel devices")
			{
				while($lineTF = <SourceFile>)
				{
					chomp($lineTF);
					$lineTF =~ s/^ +//;
					$lineTF =~ s/( +)/ /g;
					last if ($lineTF =~ "end on failure");
					#print $lineTF,"\n";
					next if (length($lineTF) < 2);
					my @Vlist = split('\ ', $lineTF);
					#print substr($Vlist[2],-2,1),substr($value,-1,1),"\n";

					if(substr($value,-1,1) eq "m" and substr($Vlist[2],-2,1) eq "m") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2))."m";}
					if(substr($value,-1,1) eq "m" and substr($Vlist[2],-2,1) eq "u") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)/1000)."m";}

					if(substr($value,-1,1) eq "u" and substr($Vlist[2],-2,1) eq "u") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2))."u";}
					if(substr($value,-1,1) eq "u" and substr($Vlist[2],-2,1) eq "m") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000)."u";}
					if(substr($value,-1,1) eq "u" and substr($Vlist[2],-2,1) eq "n") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)/1000)."u";}

					if(substr($value,-1,1) eq "n" and substr($Vlist[2],-2,1) eq "n") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2))."n";}
					if(substr($value,-1,1) eq "n" and substr($Vlist[2],-2,1) eq "m") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000000)."n";}
					if(substr($value,-1,1) eq "n" and substr($Vlist[2],-2,1) eq "u") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000)."n";}
					if(substr($value,-1,1) eq "n" and substr($Vlist[2],-2,1) eq "p") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)/1000)."n";}

					if(substr($value,-1,1) eq "p" and substr($Vlist[2],-2,1) eq "p") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2))."p";}
					if(substr($value,-1,1) eq "p" and substr($Vlist[2],-2,1) eq "u") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000000)."p";}
					if(substr($value,-1,1) eq "p" and substr($Vlist[2],-2,1) eq "n") {$value = (substr($value,0,-1) + substr($Vlist[2],0,-2)*1000)."p";}

					if(substr($value,-1,1) =~ '\d' and substr($Vlist[2],-2,1) =~ '\d') {$value = ($value * substr($Vlist[2],0,-1))/($value + substr($Vlist[2],0,-1));}		# ok
					if(substr($value,-1,1) =~ '\d' and substr($Vlist[2],-2,1) eq "k") {$value = ($value * substr($Vlist[2],0,-2)*1000)/($value + substr($Vlist[2],0,-2)*1000);}	#ok
					if(substr($value,-1,1) eq "k" and substr($Vlist[2],-2,1) =~ '\d') {$value = ((substr($value,0,-1)*1000) * substr($Vlist[2],0,-1))/((substr($value,0,-1)*1000) + substr($Vlist[2],0,-1));}	#ok
					
					if(substr($value,-1,1) eq "k" and substr($Vlist[2],-2,1) eq "k") {$value = (substr($value,0,-1) * substr($Vlist[2],0,-2))/(substr($value,0,-1) + substr($Vlist[2],0,-2))."k";}	#ok
					if(substr($value,-1,1) eq "k" and substr($Vlist[2],-2,1) eq "M") {$value = (substr($value,0,-1) * substr($Vlist[2],0,-2)*1000)/(substr($value,0,-1) + substr($Vlist[2],0,-2)*1000)."k";}	#ok
					if(substr($value,-1,1) eq "M" and substr($Vlist[2],-2,1) eq "k") {$value = ((substr($value,0,-1)*1000) * substr($Vlist[2],0,-2))/((substr($value,0,-1)*1000) + substr($Vlist[2],0,-2))."k";}

					if(substr($value,-1,1) eq "M" and substr($Vlist[2],-2,1) eq "M") {$value = (substr($value,0,-1) * substr($Vlist[2],0,-2))/(substr($value,0,-1) + substr($Vlist[2],0,-2))."M";}
					}
				if(substr($value,0,-1) >= 1000 and substr($value,-1,1) eq "p"){$value = sprintf("%.3f",(substr($value,0,-1)/1000))."n";}
				if(substr($value,0,-1) >= 1000 and substr($value,-1,1) eq "n"){$value = sprintf("%.3f",(substr($value,0,-1)/1000))."u";}
				if(substr($value,0,-1) >= 1000 and substr($value,-1,1) eq "u"){$value = sprintf("%.3f",(substr($value,0,-1)/1000))."m";}
				
				if(substr($value,-1,1) =~ '\d' and $value < 1000){$value = sprintf("%.3f",$value);}			#ok
				if(substr($value,-1,1) eq "k" and $value =~ /\./){$value = sprintf("%.3f",substr($value,0,-1)); $value =~ s/\.?0+$//; $value = $value."k";}	#ok
				if(substr($value,-1,1) eq "M" and $value =~ /\./){$value = sprintf("%.3f",substr($value,0,-1)); $value =~ s/\.?0+$//; $value = $value."M";}	#ok

				while($value){if($value =~ '\.' and substr($value,-1,1) =~ '\d' and substr($value,-1,1) eq '0' ) {$value =  substr($value,0,-1); } else{last;}}
				while($value){if($value =~ '\.' and substr($value,-1,1) =~ '\D' and substr($value,-2,1) eq '0' ) {$value =  substr($value,0,-2).substr($value,-1,1); } elsif(substr($value,-2,1) =~ '\.' and substr($value,-1,1) =~ '\D' ) {$value =  substr($value,0,-2).substr($value,-1,1); } else{last;}}
				}

			if ($list[0] =~ /^(resistor|capacitor|inductor|zener)$/)							## matching type ##
			{	
				$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
				$tested-> write($rowT, 1, $list[0], $format_data);  							## TestType ##
				$tested-> write($rowT, 2, $value, $format_data);  								## BOMvalue ##

				$tested->conditional_formatting($rowT, 3,
				{
					type     => 'cell',
					criteria => 'not equal to',
					value    => $value,
					format   => $format_VCC,
					});

				@param =  split('\,', $list[1]);	
				while($param[0]){if($param[0] =~ '\.' and substr($param[0],-1,1) =~ '\d' and substr($param[0],-1,1) eq '0' ) {$param[0] =  substr($param[0],0,-1); } else{last;}}
				while($param[0]){if($param[0] =~ '\.' and substr($param[0],-1,1) =~ '\D' and substr($param[0],-2,1) eq '0' ) {$param[0] =  substr($param[0],0,-2).substr($param[0],-1,1); } elsif(substr($param[0],-2,1) =~ '\.' and substr($param[0],-1,1) =~ '\D' ) {$param[0] =  substr($param[0],0,-2).substr($param[0],-1,1); } else{last;}}
				$tested-> write($rowT, 3, $param[0], $format_data);  							## Nominal ## 
				if($param[1] <= 40){$tested-> write($rowT, 4, $param[1], $format_data);}  		## HiLimit ##
				if($param[1] > 40){$tested-> write($rowT, 4, $param[1], $format_STP);}  		## HiLimit ##
				if($param[2] <= 40){$tested-> write($rowT, 5, $param[2], $format_data);}  		## LoLimit ##
				if($param[2] > 40){$tested-> write($rowT, 5, $param[2], $format_STP);}			## LoLimit ##
				$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);  		## Comment ##
				$rowT++;	
				goto Next_Ori1;
				}	
			elsif ($list[0] eq "diode")															## matching type ##
			{	
				$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
				$tested-> write($rowT, 1, $list[0], $format_data);  							## TestType ##
				@param =  split('\,', $list[1]);
				if ($lineTF !~ m/\"/g)
				{
					$tested-> write($rowT, 4, $param[0], $format_data);  						## HiLimit ##
					$tested-> write($rowT, 5, $param[1], $format_data);	  						## LoLimit ##
					$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
					}
				if ($lineTF =~ m/\"/g)
				{
					$DioNom = "";	$DioNom =  $DioNom . $param[0];  							## Nominal ## 
					$DioHiL = "";	$DioHiL =  $DioHiL . $param[1];  							## HiLimit ##
					$DioLoL = "";	$DioLoL =  $DioLoL . $param[2];	  							## LoLimit ##
					$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
					while($lineTF = <SourceFile>)
					{
						chomp($lineTF);
						$lineTF =~ s/^ +//;
						if (substr($lineTF,0,5) eq "diode")
						{
							#print $lineTF."\n";
							@param =  split('\,', substr($lineTF,6));
							$DioNom =  $DioNom . "\n" . $param[0];  							## Nominal ## 
							$DioHiL =  $DioHiL . "\n" . $param[1];  							## HiLimit ##
							$DioLoL =  $DioLoL . "\n" . $param[2];	  							## LoLimit ##
							}
						elsif (eof)
						{
							$tested-> write($rowT, 3, $DioNom, $format_data);  					## Nominal ## 
							$tested-> write($rowT, 4, $DioHiL, $format_data);  					## HiLimit ##
							$tested-> write($rowT, 5, $DioLoL, $format_data);	  				## LoLimit ##
							last;
							}
						}
					}
				$rowT++;
				goto Next_Ori1;
				}
			elsif ($list[0] eq "jumper")														## matching type ##
			{
				$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
				$tested-> write($rowT, 1, $list[0], $format_data);  							## TestType ##
				if($value eq '-'){$tested-> write($rowT, 2, $value, $format_data);}  			## BOMvalue ##
				elsif((substr($value,-1,1) =~ '\d') and $value < 10){$tested-> write($rowT, 2, $value, $format_data);}
				else{$tested-> write($rowT, 2, $value, $format_VCC);}
				@param =  split('\,', $list[1]);
				$OP = 0; if ($lineTF =~ "op") {$OP = 1;}
				if($OP == 0){$tested-> write($rowT, 4, $param[0], $format_data);$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);}  ## Excel ##
				if($OP == 1){$tested-> write($rowT, 5, $param[0], $format_data);$tested-> write($rowT, 6, "OP test", $format_STP);}	## Excel ##
				$rowT++;
				goto Next_Ori1;
				}
			elsif ($list[0] eq "fuse")					  										## matching type ##
			{
				$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
				$tested-> write($rowT, 1, $list[0], $format_data);  							## TestType ##
				@param =  split('\,', $list[1]);
				$tested-> write($rowT, 4, $param[0], $format_data);
				$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);  		## Excel ##
				$rowT++;
				goto Next_Ori1;
				}
			elsif (eof and $foundTO == 1)					  									## no parameter ##
			{
				$tested-> write($rowT, 0, $Mult_file, $format_item);  								## Excel ##
				$tested-> write($rowT, 1, "[base ver]", $format_data);  							## TestType ##
				$tested-> write($rowT, 6, "No Test Parameter Found in TestFile.", $format_anno1);  	## Excel ##
				$rowT++;
				goto Next_Ori1;
				}
			}}
		Next_Ori1:
	#%%%%%%%%%%%%%%% testable device [sub-version] %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	for ($v = 0; $v < $len_ver; $v = $v + 1)
	{
		if (exists($testorder{$versions[$v]."+".$Mult_file}))
		{
			if(substr($testplan{$Mult_file},0,6) eq "tested"
			and (grep{ $_ eq $testorder{$versions[$v]."+".$Mult_file}} @TOtestedver))
			{
				$foundTO = 1;
				print "			Multiple AnaVTest	", $Mult_file." - [$versions[$v]]\n";
				@array = ("-","-","-","-","-");
				$array_ref = \@array;
				$tested-> write_row($rowT, 2, $array_ref, $format_data);
				
				$UNCover = 1;	
				$coverage-> write($rowC, 1, 'V', $format_data);			#Coverage
	
				open(SourceFile, "<$versions[$v]/analog/$Mult_file") or open(SourceFile, "<$versions[$v]/analog/1%$Mult_file") or warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
				if ($! eq "No such file or directory"){$tested-> write($rowT, 6, "TestFile not found.", $format_anno1);}
				else
				{
				while($lineTF = <SourceFile>)							#read parameter
				{
					$len = 0;
					chomp($lineTF);
					$lineTF =~ s/^ +//;                               	#clear head of line spacing
					#print $lineTF,"\n";
					next if (length($lineTF) < 2);
					my @list = split('\ ', $lineTF, 2);
					#print $list[0],"\n";
					#print $list[1],"\n";
					
					if ($list[0] =~ /^(resistor|capacitor|inductor|zener)$/)							## matching type ##
					{
						$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ".$list[0], $format_data);  	## TestType ##
						@param =  split('\,', $list[1]);
						$tested-> write($rowT, 3, $param[0], $format_data);  							## Nominal ## 
						if($param[1] <= 40){$tested-> write($rowT, 4, $param[1], $format_data);}  		## HiLimit ##
						if($param[1] > 40){$tested-> write($rowT, 4, $param[1], $format_STP);}  		## HiLimit ##
						if($param[2] <= 40){$tested-> write($rowT, 5, $param[2], $format_data);}  		## LoLimit ##
						if($param[2] > 40){$tested-> write($rowT, 5, $param[2], $format_STP);}			## LoLimit ##
						$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);  		## Comment ##
						$rowT++;
						goto Next_Rev1;
						}
					elsif ($list[0] eq "diode")															## matching type ##
					{
						$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ".$list[0], $format_data);  	## TestType ##
						@param =  split('\,', $list[1]);
						if ($lineTF !~ m/\"/g)
						{
							$tested-> write($rowT, 4, $param[0], $format_data);  						## HiLimit ##
							$tested-> write($rowT, 5, $param[1], $format_data);	  						## LoLimit ##
							$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
							}
						if ($lineTF =~ m/\"/g)
						{
							$DioNom = "";	$DioNom =  $DioNom . $param[0];  							## Nominal ## 
							$DioHiL = "";	$DioHiL =  $DioHiL . $param[1];  							## HiLimit ##
							$DioLoL = "";	$DioLoL =  $DioLoL . $param[2];	  							## LoLimit ##
							$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);  	## Comment ##
							while($lineTF = <SourceFile>)
							{
								chomp($lineTF);
								$lineTF =~ s/^ +//;
								if (substr($lineTF,0,5) eq "diode")
								{
								#print $lineTF."\n";
								@param =  split('\,', substr($lineTF,6));
									$DioNom =  $DioNom . "\n" . $param[0];  							## Nominal ## 
									$DioHiL =  $DioHiL . "\n" . $param[1];  							## HiLimit ##
									$DioLoL =  $DioLoL . "\n" . $param[2];	  							## LoLimit ##
									}
								elsif (eof)
								{
									$tested-> write($rowT, 3, $DioNom, $format_data);  					## Nominal ## 
									$tested-> write($rowT, 4, $DioHiL, $format_data);  					## HiLimit ##
									$tested-> write($rowT, 5, $DioLoL, $format_data);	  				## LoLimit ##
									last;
									}
								}
							}
						$rowT++;
						goto Next_Rev1;
						}
					elsif ($list[0] eq "jumper")														## matching type ##
					{
						$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ".$list[0], $format_data);  	## TestType ##
						@param =  split('\,', $list[1]);
						$OP = 0; if ($lineTF =~ "op") {$OP = 1;}
						if($OP == 0){$tested-> write($rowT, 4, $param[0], $format_data);$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);}  ## Excel ##
						if($OP == 1){$tested-> write($rowT, 5, $param[0], $format_data);$tested-> write($rowT, 6, "OP test", $format_STP);}	## Excel ##
						$rowT++;
						goto Next_Rev1;
						}
					elsif ($list[0] eq "fuse")					  										## matching type ##
					{
						$tested-> write($rowT, 0, $Mult_file, $format_item);  							## TestName ##
						$tested-> write($rowT, 1, "[".$versions[$v]."] - ".$list[0], $format_data);  	## TestType ##
						@param =  split('\,', $list[1]);
						$tested-> write($rowT, 4, $param[0], $format_data);
						$tested-> write($rowT, 6, substr($testplan{$Mult_file},6), $format_anno);  		## Excel ##
						$rowT++;
						goto Next_Rev1;
						}
					elsif (eof and $foundTO == 1)					  									## no parameter ##
					{
						$tested-> write($rowT, 0, "[".$versions[$v]."] - ", $format_STP);  					## Excel ##
						$tested-> write($rowT, 6, "No Test Parameter Found in TestFile.", $format_anno1);  	## Excel ##
						$rowT++;
						goto Next_Rev1;
						}
					}}
				Next_Rev1:
				}
		#%%%%%%%%%%%%%%% testorder skipped devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
			elsif(grep{ $_ eq $testorder{$versions[$v]."+".$Mult_file}} @TOuntestver)
			{
				$foundTO = 1;
				print "			Multiple NullTest	", $Mult_file." - [$versions[$v]]\n";
				$tested-> write($rowT, 0, $Mult_file, $format_data);  										## Excel ##
				$tested-> write($rowT, 1, "[".$versions[$v]."] - ", $format_STP);  							## TestType ##
				@array = ("-","-","-","-");
					$array_ref = \@array;
					$tested-> write_row($rowT, 2, $array_ref, $format_data);
				$tested-> write($rowT, 6, "test is skipped in version [$versions[$v]]", $format_anno1);  	## Excel ##
				$rowT++;
				#last;
				}
			}
		}
	if ($rowT - $row_ver > 1){$tested-> merge_range($row_ver, 0, $rowT-1, 0, $Mult_file, $format_item);}
	}
	#%%%%%%%%%%%%%%% testplan skipped devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif (exists($testplan{$Mult_file})
	and substr($testplan{$Mult_file},0,7) eq "skipped"
	and grep{ $_ eq $testorder{$Mult_file}} @TOtested)
	{
		$foundTO = 1;
		print "			Multiple SkipTest	", $Mult_file,"\n";
		if ($UNCover == 0){$coverage-> write($rowC, 1, 'K', $format_VCC);}									#Coverage
		$untest-> write($rowU, 0, $Mult_file, $format_data);												## Excel ##
		$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_anno1);  							## Excel ##
		$untest-> write($rowU, 2, substr($testplan{$Mult_file},8), $format_anno);  							## Excel ##
		$rowU++;
		#last; #goto Next_Dev;
		}
	#%%%%%%%%%%%%%%% testplan inexistent devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif (not exists($testplan{$Mult_file})
	and grep{ $_ eq $testorder{$Mult_file}} @TOtested)
	{
		$foundTO = 1;
		print "			Testplan InExistent	", $Mult_file,"\n";
		if ($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_data);}									#Coverage
		$untest-> write($rowU, 0, $Mult_file, $format_data);												## Excel ##
		$untest-> write($rowU, 1, "Unidentified in Testplan.", $format_VCC);
		$untest-> write($rowU, 2, $testorder{$Mult_file}, $format_anno);									## Excel ##
		$rowU++;
		#last; #goto Next_Dev;
		}
	#%%%%%%%%%%%%%%% testorder NullTest devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif(exists($testorder{$Mult_file})
	and grep{ $_ eq $testorder{$Mult_file}} @TOuntest)
	{
		$foundTO = 1;
		print "			Multiple NullTest	", $Mult_file,"\n";
		$untest-> write($rowU, 0, $Mult_file, $format_data);												## Excel ##
		$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_anno);							## Excel ##
		$UTline = "";
		$fileF = 0;
		
		if($UNCover == 0){$coverage-> write($rowC, 1, 'N', $format_data);}									#Coverage
		open(ALL, "<analog/$Mult_file") or open(ALL, "<analog/1%$Mult_file") or warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
		if ($! eq "No such file or directory"){$untest-> write($rowU, 2, "TestFile not found.", $format_anno1);}
		else
		{
		while($line = <ALL>)
		{
			$fileF = 1;
			if (index($line,$device)>1)
			{
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;
				}
			elsif (eof){last;}
			}}
		$UTline =~ s/(^\s+|\s+$)//g;
		if($UTline eq "" and $fileF == 1){$untest-> write($rowU, 2, "No Comments Found in TestFile.", $format_anno1);}
		if($UTline ne "")
		{
			$untest-> write($rowU, 2, $UTline, $format_anno);
			$untest-> set_column(2, 2, $length_anno);
			}
		$rowU++;
		close ALL;
		#last;
		}
	#%%%%%%%%%%%%%%% testorder skipped devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif(exists($testorder{$Mult_file})
	and grep{ $_ eq $testorder{$Mult_file}} @TOskipped)
	{
		$foundTO = 1;
		print "			Multiple Skipped	", $Mult_file,"\n";
		$untest-> write($rowU, 0, $Mult_file, $format_data);													## Excel ##
		$untest-> write($rowU, 1, "been Skipped in TestOrder.", $format_anno1);									## Excel ##
		$UTline = "";
		$fileF = 0;
		
		if($UNCover == 0){$coverage-> write($rowC, 1, 'K', $format_VCC);}										#Coverage
		open(ALL, "<analog/$Mult_file") or open(ALL, "<analog/1%$Mult_file") or warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
		if ($! eq "No such file or directory"){$untest-> write($rowU, 2, "TestFile not found.", $format_anno1);}
		else
		{
		while($line = <ALL>)
		{
			$fileF = 1;
			if (index($line,$device)>1)
			{
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;
				}
			elsif (eof){last;}
			}}
		$UTline =~ s/(^\s+|\s+$)//g;
		if($UTline eq "" and $fileF == 1){$untest-> write($rowU, 2, "No Comments Found in TestFile.", $format_anno1);}
		if($UTline ne "")
		{
			$untest-> write($rowU, 2, $UTline, $format_anno);
			$untest-> set_column(2, 2, $length_anno);
			}
		$rowU++;
		close ALL;
		#next; #goto Next_Dev;
		}
	#%%%%%%%%%%%%%%% parallel tested devices %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif(exists($testorder{$Mult_file})
	and grep{ $_ eq $testorder{$Mult_file}} @TOparal)
	{
		$foundTO = 1;
		print "			Multiple ParalTest	", $Mult_file,"\n";
		
		if($UNCover == 0){$coverage-> write($rowC, 1, 'L', $format_data);}					#Coverage
		$limited-> write($rowL, 0, $Mult_file, $format_data);								## Excel ##
		$UTline = "";
		$fileF = 0;
	
		open(ALL, "<analog/$Mult_file") or open(ALL, "<analog/1%$Mult_file") or warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
		if ($! eq "No such file or directory"){$limited-> write($rowL, 1, "TestFile not found.", $format_anno1);}
		else
		{
		while($line = <ALL>)
		{
			$fileF = 1;
			if (index($line,$device)>1)
			{
				$line = substr($line,1);
				$line =~ s/(^\s+)//g;
				if (length($line)> $length_anno){$length_anno = length($line);}
				$UTline = $line . $UTline;
				}
			elsif (eof){last;}
			}}
		$UTline =~ s/(^\s+|\s+$)//g;
		if($UTline eq "" and $fileF == 1){$limited-> write($rowL, 1, "No Comments Found in TestFile.", $format_anno1);}
		if($UTline ne "")
		{
			$limited-> write($rowL, 1, $UTline, $format_anno);	
			$limited-> set_column(1, 1, $length_anno);
			}
		$rowL++;
		close ALL;
		#last;
		}
	#%%%%%%%%%%%%%%% testable analog powered test %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif(exists($testorder{$Mult_file})
	and ($testorder{$Mult_file} eq "tested-pwr" or $testorder{$Mult_file} eq "untest-pwr"
	or $testorder{$Mult_file} eq "tested-dio" or $testorder{$Mult_file} eq "untest-dio")
    )
	{
		$foundTO = 1;
		print "			Multiple PwrTest	", $Mult_file,"\n";
			@array = ("-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 3, $array_ref, $format_data);

		if ($testorder{$Mult_file} eq "tested-pwr")
		{
			if(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,6) eq "tested")
			{
				$cover = 1;
				$coverage-> write($rowC, 3, 'V', $format_data);													#Coverage
				$power-> write($rowP, 0, $Mult_file, $format_data);												## Excel ##
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);									## Excel ##
				}
			elsif(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,7) eq "skipped")
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'K', $format_VCC);}								#Coverage
				$power-> write($rowP, 0, $Mult_file, $format_data);												## Excel ##
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);			## Excel ##
				}
			elsif(not exists($testplan{$Mult_file}))
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}								#Coverage
				$power-> write($rowP, 0, $Mult_file, $format_data);	
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);
				$power-> write($rowP, 2, "Unidentified in testplan.", $format_VCC);
				}
				
		if (exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,6) eq "tested")
		{
			open(ALL, "<analog/$Mult_file") or open(ALL, "<analog/1%$Mult_file")  or  warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
			if ($! eq "No such file or directory"){$power-> write($rowP, 3, "TestFile not found.", $format_anno1);}
			else
			{
				@testname = ();
				while($line = <ALL>)
				{
					chomp ($line);
					$line =~ s/(^\s+|\s+$)//g;
					next if ($line eq "");
					#print $line,"\n";
					my @list = split('\"', $line);
					$list[0] =~ s/(^\s+|\s+$)//g;
					if ($list[0] eq "test")
					{
						#print $list[1],"\n";
						push(@testname, uc($list[1]));
						}
					last if ($line =~ "end test")
					}
				while($line = <ALL>)
				{
					chomp ($line);
					$line =~ s/(^\s+|\s+$)//g;
					next if ($line eq "");
					my @list = split('\"', $line);
					$list[0] =~ s/(^\s+|\s+$)//g;
					if ($list[0] eq "subtest")
					{
						foreach my $i (0..@testname-1)
						{
							if (grep{ $_ eq uc($list[1])} @testname and uc($list[1]) eq uc($testname[$i]))
							{
								while($line = <ALL>)
								{
									$line =~ s/(^\s+|\s+$)//g;
									if (substr($line,0,7) eq "measure")
									{
										#print $testname[$i],"\n";
										$testname[$i]= $testname[$i]." / ".$line."\n";
										#print $testname[$i],"\n";
										goto OUTER;
										}
									}
								}
							}
						}
					OUTER:
					}
				}
			close ALL;
			$content = join("",@testname);
			$content =~ s/(^\s+|\s+$)//g;
			$power-> write($rowP, 3, $content, $format_anno);
			if($content eq ""){$power-> write($rowP, 3, "NO valid test found.", $format_anno1);}
			}
		$rowP++;
		}

		elsif($testorder{$Mult_file} eq "untest-pwr")
		{
			if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}				#Coverage
			@array = ("-","-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 2, $array_ref, $format_data);

			$power-> write($rowP, 0, $Mult_file, $format_data);	
			$power-> write($rowP, 1, "Skipped - $Mult_file", $format_anno1);  				## Excel ##
			$rowP++;
			}
		}
	#%%%%%%%%%%%%%%% testable digital test %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif(exists($testorder{$Mult_file})
	and ($testorder{$Mult_file} eq "tested-dig" or $testorder{$Mult_file} eq "untest-dig")
    )
	{
		$foundTO = 1;
		$length_DigPin = 10;
		print "			Multiple DigiTest	", $Mult_file,"\n";
			@array = ("-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 3, $array_ref, $format_data);

		if ($testorder{$Mult_file} eq "tested-dig")
		{
			if(exists($testplan{$Mult_file})and substr($testplan{$Mult_file},0,6) eq "tested")
			{
				$cover = 1;
				$coverage-> write($rowC, 2, 'V', $format_data);													#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);									## Excel ##
				}
			elsif(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,7) eq "skipped")
			{
				if ($cover == 0){$coverage-> write($rowC, 2, 'K', $format_VCC);}								#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);			## Excel ##
				}
			elsif(not exists($testplan{$Mult_file}))
			{
				if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_data);}								#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);
				$power-> write($rowP, 2, "Unidentified in testplan.", $format_VCC);
				}

			###### hyperlink #############################################################
			$power-> write_url($rowP, 0, 'internal:'.$device.'!A1');		## hyperlink
			
			if ($worksheet == 0)
			{
				$worksheet = 1;
				$IC = $bom_coverage_report-> add_worksheet($device);		## hyperlink
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
							#print $lineDig."\n";
							if ($lineDig eq uc($device))
							{
								while($lineDig = <Boards>) 
								{
									#print $lineDig;
									$Total_Pin++;
									@DigPin = split('\.',$lineDig);
									$DigPin[0] =~ s/(^\s+|\s+$)//g;
									$DigPin[1] =~ s/(^\s+|\s+$)//g;
									#print $DigPin[0]."\n";
									if ($DigPin[1] =~ /(GND|GROUND)/)
									{
										if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_GND); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
										if ($DigPin[0] =~ /^\D/i)
										{
											$IC-> write($DigPin[0], $DigPin[1], $format_GND);
											($pos) = $DigPin[0] =~ /^\D+/g;
											if (length($pos) == 1){$DigPos = ord($pos)%64;}
											if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
											#print $DigPos."\n";
											if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
											}
										$GND_Pin++;
										}
									elsif ($DigPin[1] =~ /^(\+0|0V|\+1|1V|\+2|2V|\+3|3V|\+5|5V|V_|VCC|VDD|PP|P0V|P1V|P2V|P3V|P5V)/)
									{
										if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_VCC); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
										if ($DigPin[0] =~ /^\D/i)
										{
											$IC-> write($DigPin[0], $DigPin[1], $format_VCC);
											($pos) = $DigPin[0] =~ /^\D+/g;
											if (length($pos) == 1){$DigPos = ord($pos)%64;}
											if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
											#print $DigPos."\n";
											if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
											}
										$Power_Pin++;
										}
									elsif ($DigPin[1] =~ /(^NC_|_NC$|NONE)/)
									{
										if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1], $format_NC); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
										if ($DigPin[0] =~ /^\D/i)
										{
											$IC-> write($DigPin[0], $DigPin[1], $format_NC);
											($pos) = $DigPin[0] =~ /^\D+/g;
											if (length($pos) == 1){$DigPos = ord($pos)%64;}
											if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
											#print $DigPos."\n";
											if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
											}
										$NC_Pin++;
										}
									else{
										if (exists($bdg_list{uc($device)."\.".$DigPin[0]}))
										{
											#print uc($device)."\.".$DigPin[0],"\n";
											if($bdg_list{uc($device)."\.".$DigPin[0]})
											{
												#print uc($device)."\.".$DigPin[0],"\n";
												if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Toggle_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
												if ($DigPin[0] =~ /^\D/i)
												{
													$IC-> write($DigPin[0], $DigPin[1]."\n* Toggle_Test", $format_data);
													($pos) = $DigPin[0] =~ /^\D+/g;
													if (length($pos) == 1){$DigPos = ord($pos)%64;}
													if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
													if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
													}
												}
											}
										else{
											if(exists($hash_pin{$DigPin[1]}))
											{
												if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1]."\n* Contact_Test", $format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
												if ($DigPin[0] =~ /^\D/i)
												{
													$IC-> write($DigPin[0], $DigPin[1]."\n* Contact_Test", $format_data);
													($pos) = $DigPin[0] =~ /^\D+/g;
													if (length($pos) == 1){$DigPos = ord($pos)%64;}
													if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
													#print $DigPin[1]."\n";
													if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
													}
												}
											else{
											if ($DigPin[0] =~ /^\d/){$IC-> write(int($DigPin[0])-1, 0, $DigPin[1],$format_data); if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column(0, 0, $length_DigPin+2);}
											if ($DigPin[0] =~ /^\D/i)
											{
												$IC-> write($DigPin[0], $DigPin[1], $format_data);
												($pos) = $DigPin[0] =~ /^\D+/g;
												if (length($pos) == 1){$DigPos = ord($pos)%64;}
												if (length($pos) == 2){$DigPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
												if (length($DigPin[1])> $length_DigPin){$length_DigPin = length($DigPin[1]);} $IC-> set_column($DigPos-1, $DigPos-1, $length_DigPin+2);
												}}
											}
										}
									if ($lineDig =~ "\;")
									{
										$power-> write($rowP, 4, $Total_Pin, $format_item);
										$power-> write($rowP, 5, $Power_Pin, $format_VCC);
										$power-> write($rowP, 6, $GND_Pin, $format_GND);
										$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "*Toggle_Test*")', $format_data);
										$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "*Contact_Test*")', $format_data);
										$power-> write($rowP, 9, $NC_Pin, $format_NC);
										$power-> write_formula($rowP, 10, "=(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-H".($rowP+1)."-I".($rowP+1)."-J".($rowP+1).")", $format_data);
										$power-> write_formula($rowP, 11, "=(H".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
										$power-> write_formula($rowP, 12, "=(I".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
										last;
										}
									}
								}
							}
						}
  					}
  				close Boards;
  				}

			###### hyperlink #############################################################
			$power-> write($rowP, 0, $device, $format_hylk);

			$family = "";
  			open(SourceFile, "<digital/$Mult_file") or open(SourceFile, "<digital/1%$Mult_file") or  warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
			if ($! eq "No such file or directory"){$power-> write($rowP, 3, "TestFile not found.", $format_anno1);}
			else
			{
				while($lineTF = <SourceFile>)							#reading family
				{
					$lineTF =~ s/(^\s+)//g;								#clear head of line spacing
					#print $lineTF;
					if (substr($lineTF,0,6) eq "family")
					{$family = $lineTF . $family;}
  					}
  					close SourceFile;
  					chomp($family);
  					$power-> write($rowP, 3, $family, $format_anno);
  					if($family eq ""){$power-> write($rowP, 3, "Family not define.", $format_anno1);}
				}
				$rowP++;
			}

		elsif($testorder{$Mult_file} eq "untest-dig")
		{
			if ($cover == 0){$coverage-> write($rowC, 2, 'N', $format_data);}				#Coverage
			@array = ("-","-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 2, $array_ref, $format_data);

			$power-> write($rowP, 1, "Skipped - $Mult_file", $format_anno1);  				## Excel ##
			$rowP++;
			}
		}
	#%%%%%%%%%%%%%%% testable mixed device %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif(exists($testorder{$Mult_file})
	and ($testorder{$Mult_file} eq "tested-mix" or $testorder{$Mult_file} eq "untest-mix")
    )
	{
		$foundTO = 1;
		print "			Multiple MixTest	", $Mult_file,"\n";
			@array = ("-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 3, $array_ref, $format_data);

		if($testorder{$Mult_file} eq "tested-mix")
		{
			if(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,6) eq "tested")
			{
				$cover = 1;
				$coverage-> write($rowC, 3, 'V', $format_data);											#Coverage
				$power-> write($rowP, 0, $Mult_file, $format_data);										## Excel ##
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);		## Excel ##
				$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);							## Excel ##
				}
			elsif(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,7) eq "skipped")
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'K', $format_VCC);}						#Coverage
				$power-> write($rowP, 0, $Mult_file, $format_data);										## Excel ##
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);		## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);	## Excel ##
				}
			elsif(not exists($testplan{$Mult_file}))
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}						# Coverage
				$power-> write($rowP, 0, $Mult_file, $format_data);	
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);
				$power-> write($rowP, 2, "Unidentified in testplan.", $format_VCC);
				}
			$rowP++;
			}
		
		elsif($testorder{$Mult_file} eq "untest-mix")
		{
			if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}							#Coverage
			@array = ("-","-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 2, $array_ref, $format_data);
			$power-> write($rowP, 0, $Mult_file, $format_data);	
			$power-> write($rowP, 1, "Skipped - $Mult_file", $format_anno1);  							## Excel ##
			$rowP++;
			}
		}
	#%%%%%%%%%%%%%%% testable Bscan device %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	elsif(exists($testorder{$Mult_file})
	and ($testorder{$Mult_file} eq "tested-bscan" or $testorder{$Mult_file} eq "untest-bscan")
    )
	{
		$foundTO = 1;
		$length_SNail = 10;
		print "			Multiple BscTest	", $Mult_file,"\n";
		@array = ("-","-","-","-","-","-","-","-","-");
		$array_ref = \@array;
		$power-> write_row($rowP, 4, $array_ref, $format_data);

		if($testorder{$Mult_file} eq "tested-bscan")
		{
			if(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,6) eq "tested")
			{
				$cover = 1;
				$coverage-> write($rowC, 4, 'V', $format_data);													#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);									## Excel ##
				}
			elsif(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,7) eq "skipped")
			{
				if ($cover == 0){$coverage-> write($rowC, 4, 'K', $format_VCC);}								#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);				## Excel ##
				$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);			## Excel ##
				}
			elsif(not exists($testplan{$Mult_file}))
			{
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_data);}								#Coverage
				$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);
				$power-> write($rowP, 2, "Unidentified in testplan.", $format_VCC);
				}

			###### hyperlink #############################################################
			$power-> write_url($rowP, 0, 'internal:'.$device.'!A1');    	## hyperlink

			if ($worksheet == 0)
			{
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
						{
							while($lineDig = <Boards>) 
							{
								$Total_Pin++;
								@BscanNail = split('\.',$lineDig);
								$BscanNail[0] =~ s/(^\s+|\s+$)//g;
								$BscanNail[1] =~ s/(^\s+|\s+$)//g;
								if ($BscanNail[1] =~ /(GND|GROUND)/)
								{
									if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_GND); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
									if ($BscanNail[0] =~ /^\D/i)
									{
										$IC-> write($BscanNail[0], $BscanNail[1], $format_GND);
										($pos) = $BscanNail[0] =~ /^\D+/g;
										if (length($pos) == 1){$NailPos = ord($pos)%64;}
										if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
										if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
										}
									$GND_Pin++;
									}
								elsif ($BscanNail[1] =~ /^(\+0|0V|\+1|1V|\+2|2V|\+3|3V|\+5|5V|V_|VCC|VDD|PP|P0V|P1V|P2V|P3V|P5V)/)
								{
									if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_VCC); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
									if ($BscanNail[0] =~ /^\D/i)
									{
										$IC-> write($BscanNail[0], $BscanNail[1], $format_VCC);
										($pos) = $BscanNail[0] =~ /^\D+/g;
										if (length($pos) == 1){$NailPos = ord($pos)%64;}
										if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
										if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
										}
									$Power_Pin++;
									}
								elsif ($BscanNail[1] =~ /(^NC_|_NC$|NONE)/)
								{
									if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_NC); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
									if ($BscanNail[0] =~ /^\D/i)
									{
										$IC-> write($BscanNail[0], $BscanNail[1], $format_NC);
										($pos) = $BscanNail[0] =~ /^\D+/g;
										if (length($pos) == 1){$NailPos = ord($pos)%64;}
										if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
										if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
										}
									$NC_Pin++;
									}
								else{
									if(exists($hash_pin{$BscanNail[1]})	)
									{
										if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1]."\n* Contact_Test", $format_data); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
										if ($BscanNail[0] =~ /^\D/i)
										{
											$IC-> write($BscanNail[0], $BscanNail[1]."\n* Contact_Test", $format_data);
											($pos) = $BscanNail[0] =~ /^\D+/g;
											if (length($pos) == 1){$NailPos = ord($pos)%64;}
											if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
											if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
											}
										}
									else{
										if ($BscanNail[0] =~ /^\d/){$IC-> write(int($BscanNail[0])-1, 0, $BscanNail[1], $format_data); if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column(0, 0, $length_SNail+2);}
										if ($BscanNail[0] =~ /^\D/i)
										{
											$IC-> write($BscanNail[0], $BscanNail[1], $format_data);
											($pos) = $BscanNail[0] =~ /^\D+/g;
											if (length($pos) == 1){$NailPos = ord($pos)%64;}
											if (length($pos) == 2){$NailPos = int(ord(substr($pos,0,1))%64) * 26 + ord(substr($pos,1,1))%64;}
											if (length($BscanNail[1])> $length_SNail){$length_SNail = length($BscanNail[1]);} $IC-> set_column($NailPos-1, $NailPos-1, $length_SNail+2);
											}
										}
									}
								if ($lineDig =~ "\;")
								{
									$power-> write($rowP, 4, $Total_Pin, $format_item);
									$power-> write($rowP, 5, $Power_Pin, $format_VCC);
									$power-> write($rowP, 6, $GND_Pin, $format_GND);
									$power-> write_formula($rowP, 7, '=COUNTIF('.$device.'!A1:GR999, "*Toggle_Test*")', $format_data);
									$power-> write_formula($rowP, 8, '=COUNTIF('.$device.'!A1:GR999, "*Contact_Test*")', $format_data);
									$power-> write($rowP, 9, $NC_Pin, $format_NC);
									$power-> write_formula($rowP, 10, "=(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-H".($rowP+1)."-I".($rowP+1)."-J".($rowP+1).")", $format_data);
									$power-> write_formula($rowP, 11, "=(H".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
									$power-> write_formula($rowP, 12, "=(I".($rowP+1)."/(E".($rowP+1)."-F".($rowP+1)."-G".($rowP+1)."-J".($rowP+1)."))", $format_FPY);
									last;
									}
								}
							}
						}
					}
  				}
  			close Boards;
  			}

			###### hyperlink #############################################################
			$power-> write($rowP, 0, $device, $format_hylk);  								## Excel ##
			
			$family = "";
  			open(SourceFile, "<digital/$Mult_file") or open(SourceFile, "<digital/1%$Mult_file") or warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
			if ($! eq "No such file or directory"){$power-> write($rowP, 3, "TestFile not found.", $format_anno1);}
			else
			{
				while($lineTF = <SourceFile>)							#reading family
				{
					chomp($lineTF);
					$lineTF =~ s/^ +//;                               	#clear head of line spacing
					#print $lineTF;
					if (substr($lineTF,0,6) eq "family")
					{$family = $lineTF . $family;}
  					
  					if (substr($lineTF,0,5) eq "nodes")
					{
						while($lineTF = <SourceFile>)
						{
							chomp ($lineTF);
							$lineTF =~ s/(^\s+|\s+$)//g;
							next if ($lineTF eq "");
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
  				if($family eq ""){$power-> write($rowP, 3, "Family not define.", $format_anno1);}
				}
				$rowP++;
			}

		elsif($testorder{$Mult_file} eq "untest-bscan")
		{
			if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_data);}				#Coverage
			@array = ("-","-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 2, $array_ref, $format_data);
			$power-> write($rowP, 1, "Skipped - $Mult_file", $format_anno1);  				## Excel ##
			$rowP++;
			}
		}
	#%%%%%%%%%%%%%%% Bscan [sub-version] %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	for ($v = 0; $v < $len_ver; $v = $v + 1)
	{
		if (exists($testorder{$versions[$v]."+".$Mult_file}) and substr($testorder{$versions[$v]."+".$Mult_file},7,9) eq "bscan+ver")
		{
			$foundTO = 1;
			print "			Multiple BscVTest	", $Mult_file." - [$versions[$v]]\n";
			@array = ("-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 3, $array_ref, $format_data);

			if ($testorder{$versions[$v]."+".$Mult_file} eq "tested-bscan+ver")
			{
				
				if(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,6) eq "tested")
				{
					$cover = 1;
					$coverage-> write($rowC, 4, 'V', $format_data);													#Coverage
					$power-> write($rowP, 1, "tested-bscan - ".$Mult_file." - [".$versions[$v]."]", $format_anno);	## Excel ##
					$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);									## Excel ##
					}
				elsif(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,7) eq "skipped")
				{
					if ($cover == 0){$coverage-> write($rowC, 4, 'K', $format_VCC);}								#Coverage
					$power-> write($rowP, 1, "tested-bscan - ".$Mult_file." - [".$versions[$v]."]", $format_anno);	## Excel ##
					$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);			## Excel ##
					}
				elsif(not exists($testplan{$Mult_file}))
				{
					if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_data);}								#Coverage
					$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);
					$power-> write($rowP, 2, "Unidentified in testplan.", $format_VCC);
					}
				$rowP++;
				}
				
			elsif($testorder{$versions[$v]."+".$Mult_file} eq "untest-bscan+ver")
			{
				if ($cover == 0){$coverage-> write($rowC, 4, 'N', $format_data);}									#Coverage
				$power-> write($rowP, 1, "Skipped - $Mult_file -ver: [".$versions[$v]."]", $format_anno1);			## Excel ##
				$power-> write($rowP, 2, "-", $format_data);														## Excel ##
				$rowP++;
				}
			}
		}
	#%%%%%%%%%%%%%%% analog powered [sub-version] %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	for ($v = 0; $v < $len_ver; $v = $v + 1)
	{
		if (exists($testorder{$versions[$v]."+".$Mult_file}) and substr($testorder{$versions[$v]."+".$Mult_file},7,7) eq "pwr+ver")
		{
			$foundTO = 1;
			print "			Multiple PwrVTest	", $Mult_file." - [$versions[$v]]\n";
			@array = ("-","-","-","-","-","-","-","-","-","-");
			$array_ref = \@array;
			$power-> write_row($rowP, 3, $array_ref, $format_data);

			if ($testorder{$versions[$v]."+".$Mult_file} eq "tested-pwr+ver")
			{
				if(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,6) eq "tested")
				{
					$cover = 1;
					$coverage-> write($rowC, 3, 'V', $format_data);													#Coverage
					$power-> write($rowP, 1, "tested-pwr - ".$Mult_file." - [".$versions[$v]."]", $format_anno);	## Excel ##
					$power-> write($rowP, 2, "Tested - ".$Mult_file, $format_anno);									## Excel ##
					}
				elsif(exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,7) eq "skipped")
				{
					if ($cover == 0){$coverage-> write($rowC, 3, 'K', $format_VCC);}								#Coverage
					$power-> write($rowP, 1, "tested-pwr - ".$Mult_file." - [".$versions[$v]."]", $format_anno);	## Excel ##
					$power-> write($rowP, 2, "Skipped - ".substr($testplan{$Mult_file},8), $format_anno1);			## Excel ##
					}
				elsif(not exists($testplan{$Mult_file}))
				{
					if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}								#Coverage
					$power-> write($rowP, 1, $testorder{$Mult_file}." - ".$Mult_file, $format_anno);
					$power-> write($rowP, 2, "Unidentified in testplan.", $format_VCC);
					}
			if (exists($testplan{$Mult_file}) and substr($testplan{$Mult_file},0,6) eq "tested")
			{
				open(ALL, "<$versions[$v]/analog/$Mult_file") or open(ALL, "<$versions[$v]/analog/1%$Mult_file") or  warn "\t!!! Failed to open '$Mult_file' file: $!.\n";
				if ($! eq "No such file or directory"){$power-> write($rowP, 3, "TestFile not found.", $format_anno1);}
				else
				{
					@testname = ();
					while($line = <ALL>)
					{
						chomp ($line);
						$line =~ s/(^\s+|\s+$)//g;
						next if ($line eq "");
						#print $line,"\n";
						my @list = split('\"', $line);
						$list[0] =~ s/(^\s+|\s+$)//g;
						if ($list[0] eq "test")
						{
							#print $list[1],"\n";
							push(@testname, uc($list[1]));
							}
						last if ($line =~ "end test")
						}
					while($line = <ALL>)
					{
						chomp ($line);
						$line =~ s/(^\s+|\s+$)//g;
						next if ($line eq "");
						my @list = split('\"', $line);
						$list[0] =~ s/(^\s+|\s+$)//g;
						if ($list[0] eq "subtest")
						{
							foreach my $i (0..@testname-1)
							{
								if (grep{ $_ eq uc($list[1])} @testname and uc($list[1]) eq uc($testname[$i]))
								{
									while($line = <ALL>)
									{
										$line =~ s/(^\s+|\s+$)//g;
										if (substr($line,0,7) eq "measure")
										{
											#print $testname[$i],"\n";
											$testname[$i]= $testname[$i]." / ".$line."\n";
											#print $testname[$i],"\n";
											goto OUTER;
											}
										}
									}
								}
							}
						OUTER:
						}
					close ALL;
					$content = join("",@testname);
					$content =~ s/(^\s+|\s+$)//g;
					$power-> write($rowP, 3, $content, $format_anno);
					if($content eq ""){$power-> write($rowP, 3, "NO valid test found.", $format_anno1);}
					}}
				$rowP++;
				}

			#mult-ver TO skipped
			elsif($testorder{$versions[$v]."+".$Mult_file} eq "untest-pwr+ver")
			{
				if ($cover == 0){$coverage-> write($rowC, 3, 'N', $format_data);}								#Coverage
				$power-> write($rowP, 1, "Skipped - $Mult_file -ver: [".$versions[$v]."]", $format_anno1);		## Excel ##
				$power-> write($rowP, 2, "-", $format_data);													## Excel ##
				$rowP++;
				}
			}
		}
	######################################################################################
			}
		}
	}
	################ reservation #########################################################
	if((not exists ($testorder{$device}) or $testorder{$device} eq "") and $foundTO == 0)
	{
		#print $foundTO,"--5--","\n";
		print "\t\t\tNO Valid Test Found	$device\n"; 
		$coverage-> write($rowC, 5, 'N', $format_data);											#Coverage
		$untest-> write($rowU, 0, $device, $format_data);  										## Excel ##
		$untest-> write($rowU, 1, "NO valid test items found in TestOrder.", $format_anno);		## Excel ##
		$untest-> write($rowU, 2, "Check TJ/SP/other testing.", $format_anno);  				## Excel ##
		$rowU++;
		goto Next_Dev;
		}

    ######################################################################################
	#print $testorder{$device},"\n";
	Next_Dev:
	#print $device.$rowP."	".$rowP_ori."\n";
	if ($rowP - $rowP_ori > 1){$power-> merge_range($rowP_ori, 0, $rowP-1, 0, $device, $format_hylk);}
	}


############################### read Board value #########################################
sub read_value{
	my $dev = shift;
	#print $dev,"\n";
	
	open (Boards, "< board");
	while(my $array = <Boards>)								
	{
		chomp ($array);
		last if ($array =~ "CONNECTIONS");
		$array =~ s/(^\s+|\s+$)//g;
		my @list = split('\ ', $array);
		next if (scalar @list < 5);
		if($list[0] eq $dev)
		{
			#print $list[0],"\n";
			#print $list[1],"\n";
			return $list[1];
			last;
			}
		}
	close Boards;
	}


############################### check if jumper tested in shorts #########################
sub shorts_check{
	my ($ibus, $sbus) = @_;	

	if (grep{ $_ eq $ibus."\/".$sbus or $_ eq $sbus."\/".$ibus} @shorts)
	{
		#print "$ibus + $sbus existed\n";
		return "True";
		}
	else
	{
		#print "$ibus + $sbus inexisted\n";
		return "False";
		}
	}


##########################################################################################

$bom_coverage_report->close();

print  "\n  >>> Completed ...\n";

my $end_time = time();
my $duration = $end_time - $start_time;
printf "\n  runtime: %.4f Sec\n", $duration;

END_Prog:

print "\n";
# system 'pause';
exit;

