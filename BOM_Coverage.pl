#!/usr/bin/perl
print "\n";
print "*******************************************************************************\n";
print "  Bom Coverage ckecking tool for 3070 <v4.9.3>\n";
print "  Author: Noon Chen\n";
print "  A Professional Tool for Test.\n";
print "  ",scalar localtime;
print "\n*******************************************************************************\n";
print "\n";

#########################################################################################
#print "  Checking In: ";

 system "stty -echo";
 print "  Password: ";
 chomp($Ccode = <STDIN>);
 print "\n";
 system "stty echo";
 
   if ($Ccode ne "\@testpro")
   #if ($Ccode ne "TestPro")
    	{
    		print "  >>> password Wrong!\n"; goto END_Prog;
    	}
   else
   	{
    		print "  >>> password Correct.\n\n";
   	}


############################ Excel ######################################################
use Excel::Writer::XLSX;

my $bom_coverage_report = Excel::Writer::XLSX->new('BOM_Coverage_Report.xlsx');
my $summary = $bom_coverage_report-> add_worksheet('Summary');
my $tested = $bom_coverage_report-> add_worksheet('Tested');
my $untest = $bom_coverage_report-> add_worksheet('Untest');
my $limited = $bom_coverage_report-> add_worksheet('LimitTest');
my $power = $bom_coverage_report-> add_worksheet('PowerTest');
#my $ICpin = $bom_coverage_report-> add_worksheet('IC pin coverage');
my $short_thres = $bom_coverage_report-> add_worksheet('Shorts_Thres');

$tested-> freeze_panes(1,1);			#冻结行、列
$untest-> freeze_panes(1,0);			#冻结行、列
$limited-> freeze_panes(1,0);			#冻结行、列
$short_thres-> freeze_panes(1,0);		#冻结行、列

$summary-> set_column(0,2,20);			#设置列宽
$tested-> set_column(0,5,20);			#设置列宽
$untest-> set_column(0,2,30);			#设置列宽
$limited-> set_column(0,1,30);			#设置列宽
$power-> set_column(0,2,30);			#设置列宽
$short_thres-> set_column(0,1,30);		#设置列宽

$summary-> activate();					#设置初始可见

#新建一个格式
$format_item = $bom_coverage_report-> add_format(bold=>1, align=>'left', border=>1, size=>12, bg_color=>'cyan');
$format_head = $bom_coverage_report-> add_format(bold=>1, align=>'vcenter', border=>1, size=>12, bg_color=>'lime');
$format_data = $bom_coverage_report-> add_format(align=>'center', border=>1);
$format_anno = $bom_coverage_report-> add_format(align=>'left', border=>1);
$format_PCT  = $bom_coverage_report-> add_format(align=>'center', border=>1, num_format=> '10');
$format_STP  = $bom_coverage_report-> add_format(color=>'red', align=>'center', border=>1, bg_color=>'yellow');
$format_anno -> set_text_wrap();

$row = 0; $col = 0;
$tested-> write($row, $col, '<Device>', $format_head);
$row = 0; $col = 1;
$tested-> write($row, $col, '<TYPE>', $format_head);
$row = 0; $col = 2;
$tested-> write($row, $col, '<Nominal>', $format_head);
$row = 0; $col = 3;
$tested-> write($row, $col, '<HiLimit>', $format_head);
$row = 0; $col = 4;
$tested-> write($row, $col, '<LoLimit>', $format_head);
$row = 0; $col = 5;
$tested-> write($row, $col, '<Comments>', $format_head);

$row = 0; $col = 0;
$untest-> write($row, $col, '<Device>', $format_head);
$row = 0; $col = 1;
$untest-> write($row, $col, '<Justification>', $format_head);
$row = 0; $col = 2;
$untest-> write($row, $col, '<Comments>', $format_head);

$row = 0; $col = 0;
$limited-> write($row, $col, '<Device>', $format_head);
$row = 0; $col = 1;
$limited-> write($row, $col, '<Comments>', $format_head);

$row = 0; $col = 0;
$power-> write($row, $col, '<TestOrder>', $format_head);
$row = 0; $col = 1;
$power-> write($row, $col, '<TestPlan>', $format_head);
$row = 0; $col = 2;
$power-> write($row, $col, '<Family>', $format_head);

$row = 0; $col = 0;
$short_thres-> write($row, $col, 'Nodes', $format_head);
$row = 0; $col = 1;
$short_thres-> write($row, $col, 'Threshold', $format_head);

$row = 0; $col = 0;
$summary-> write($row, $col, 'Test Items', $format_head);
$row = 0; $col = 1;
$summary-> write($row, $col, 'Quantity', $format_head);
$row = 0; $col = 2;
$summary-> write($row, $col, 'Percentage', $format_head);

$row = 1; $col = 0;
$summary-> write($row, $col, 'Tested', $format_item);
$row = 2; $col = 0;
$summary-> write($row, $col, 'Untest', $format_item);
$row = 3; $col = 0;
$summary-> write($row, $col, 'LimitTest', $format_item);
$row = 4; $col = 0;
$summary-> write($row, $col, 'PowerTest', $format_item);
$row = 5; $col = 0;
$summary-> write($row, $col, 'Node accessibility rate', $format_item);

$row = 1; $col = 1;
$summary-> write($row, $col, '=COUNTA(Tested!A2:A9999)', $format_data);
$row = 2; $col = 1;
$summary-> write($row, $col, '=COUNTA(Untest!A2:A9999)', $format_data);
$row = 3; $col = 1;
$summary-> write($row, $col, '=COUNTA(LimitTest!A2:A9999)', $format_data);
$row = 4; $col = 1;
$summary-> write($row, $col, '=COUNTA(PowerTest!A2:A9999)', $format_data);
$row = 5; $col = 1;
$summary-> write($row, $col, '=COUNTA(Shorts_Thres!A2:A9999)-COUNTIF(Shorts_Thres!A2:A9999,"!nodes *")', $format_data);

$row = 1; $col = 2;
$summary-> write_formula($row, $col, "=(B2/(B2+B3+B4+B5))", $format_PCT);  #输出Percentage
$row = 2; $col = 2;
$summary-> write_formula($row, $col, "=(B3/(B2+B3+B4+B5))", $format_PCT);  #输出Percentage
$row = 3; $col = 2;
$summary-> write_formula($row, $col, "=(B4/(B2+B3+B4+B5))", $format_PCT);  #输出Percentage
$row = 4; $col = 2;
$summary-> write_formula($row, $col, "=(B5/(B2+B3+B4+B5))", $format_PCT);  #输出Percentage
$row = 5; $col = 2;
$summary-> write_formula($row, $col, "=(B6/COUNTA(Shorts_Thres!A2:A9999))", $format_PCT);  #输出Percentage

$tested-> write("H2", 'Type', $format_item);
$tested-> write("H3", 'Count', $format_item);
$tested-> write("I2", 'Resistor', $format_head);
$tested-> write("J2", 'Capacitor', $format_head);
$tested-> write("K2", 'Inductor', $format_head);
$tested-> write("L2", 'Jumper', $format_head);
$tested-> write("M2", 'Diode', $format_head);
$tested-> write("N2", 'Zener', $format_head);
$tested-> write_formula("I3",, '=COUNTIF(B2:B99999,"Resistor")', $format_data);
$tested-> write_formula("J3", '=COUNTIF(B2:B99999,"Capacitor")', $format_data);
$tested-> write_formula("K3", '=COUNTIF(B2:B99999,"Inductor")', $format_data);
$tested-> write_formula("L3", '=COUNTIF(B2:B99999,"Jumper")', $format_data);
$tested-> write_formula("M3", '=COUNTIF(B2:B99999,"Diode")', $format_data);
$tested-> write_formula("N3", '=COUNTIF(B2:B99999,"Zener")', $format_data);

my $chart = $bom_coverage_report-> add_chart( type => 'pie', embedded => 1 );
$chart-> add_series(
    name       => '=Summary!$C$1',
    categories => '=Summary!$A$2:$A$5',
    values     => '=Summary!$B$2:$B$5',
    data_labels => {value => 1},
	);
$summary-> insert_chart('A9', $chart, 10, 0, 1.0, 1.6);

$rowT = 1;
$rowU = 1;
$rowL = 1;
$rowP = 1;


print "  please specify BOM list file: ";
   $bom=<STDIN>;
   chomp $bom;

open (BOM, "< $bom");                      #read BOM list.

 #4.4# open (Tested, ">BOM_Tested");              #testable device list.
 #4.4# printf Tested "%-30s","<Device>";printf Tested "%-20s","<TYPE>";printf Tested "%-20s","<Nominal>";printf Tested"%-16s","<HiLimit>";printf Tested"%-16s","<LoLimit>";printf Tested"%-30s","<Comment>";printf Tested"\n";

 #4.4# open (Nulltested, ">BOM_Nulltested");      #untestable device list.
 #4.4# open (Limited, ">BOM_LimitTested");        #limited test device list.
 #4.4# open (PowerTest, ">BOM_PowerTest");

print "\n";
print "  gether all BOM devices...";
@bom_list = <BOM>;
print "[DONE]\n";


foreach $device (@bom_list)
{
	$foundTO = 0;
	$foundTP = 0;
	$device =~ s/\s//g;                     #clear all spacing
	$device = lc($device);

	$len_device = length($device);
	print "	Analyzing ", $device, " .....\n";
	#************************ Testorder Checking *******************************************
	open (TesO, "<testorder");																#### check testorder ####
	$len = 0;
		while($lineTO = <TesO>)
			{
			chomp($lineTO);
			$lineTO =~ s/^ +//;									#clear head of line spacing
			undef @array;
			while ($lineTO =~ m/\"/g)
				{
					$leng = pos($lineTO);
					#print $leng,"\n";
					push(@array,$leng);
				}
					$StartBit_TO = shift(@array); 				#matching beginning
					#print $StartBit_TO,"\n";
					$Stop_TO = shift(@array); 					#matching stop
					$StopBit_TO = $Stop_TO - $StartBit_TO - 1;
					#print $StopBit_TO,"\n";
					#print substr($lineTO,$StartBit_TO,$StopBit_TO),"\n";
					#print length(substr($lineTO,$StartBit_TO,$StopBit_TO)),"\n"
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
			if(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO == -1
          and $skipTO == -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
          #or substr($lineTO,index($lineTO,$device) + length($device),1) eq "\_"
          #or substr($lineTO,index($lineTO,$device) + length($device),1) eq "\%")
				{
					$foundTO = 1;
					print "			General Test\n";   #, $lineTO,"\n";
					$testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
				  #print $testname,"\n";
					#print length($testname),"\n";
					open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching test name
								and substr($lineTP,0,1) ne "\!")							#matching not skipped test name
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}	
								#print $testfile,"\n";
									  if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
									open (SourceFile, "<$testfile");
									while($lineTF = <SourceFile>)							#read parameter
									{
										$len = 0;
										chomp;
										$lineTF =~ s/^ +//;                               	#clear head of line spacing
										#print $lineTF;
										if (substr($lineTF,0,8) eq "resistor")				#### matching resistor ########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,8);										#matching type 
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											#print $parameter,'\n';
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal value
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,9) eq "capacitor")				#### matching capacitor ########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,9);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,8) eq "inductor")					#### matching inductor ########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,8);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,5) eq "diode")					#### matching diode ######
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item); ## Excel ##
											$type = substr($lineTF,0,5);						#matching type
											#4.4# printf Tested "%-40s", $type;
											$tested-> write($rowT, 1, $type, $format_data); 	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													$tested-> write($rowT, 3, $HiLimit, $format_data);  ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													$tested-> write($rowT, 4, $LoLimit, $format_data);  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,5) eq "zener")					 ####matching zener##
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,5);						 #matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	 ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,6) eq "jumper")						####matching jumper########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,6);						 #matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	 ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
											$OP = 0; if ($lineTF =~ "op") {$OP = 1;}
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-52s", $Nominal;
													if($OP == 0)
													{$tested-> write($rowT, 3, $Nominal, $format_data);
													 $tested-> write($rowT, 5, $commentTP, $format_data);}  ## Excel ##
													if($OP == 1)
													{$tested-> write($rowT, 4, $Nominal, $format_STP);
													 $tested-> write($rowT, 5, "OP test", $format_STP);}	## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,4) eq "fuse")					  	####matching fuse########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,4);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-52s", $Nominal;
													$tested-> write($rowT, 3, $Nominal, $format_data);  	## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  				#TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);    ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (eof)					  														####no parameter######
										{
										#4.4# printf Nulltested "%-30s", $testname; print Nulltested "No Test Parameter Found in TestFile."."\n";
										$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
										$untest-> write($rowU, 1, "No Test Parameter Found in TestFile.", $format_data);  	## Excel ##
										$rowU++;
											}
  								}
								}
							elsif (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname									#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}																		#ignore duplicated test name
  								$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n";
							$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
							$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							$rowU++;
							goto Next_Dev;}
							}
					close TesP;
				}
			################ multi-test(_)unpowered testable devices ###########################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),1) eq "\_"
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
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
					print "			Multi_Test(_)UNP  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					$testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
					#print $testname,"\n";
					#print length($testname),"\n";
					open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching device name
								and substr($lineTP,0,1) ne "\!")							#matching not skipped test item
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}	
								#print $testfile,"\n";
									  if ($testfile eq $testfile_last){last;}																	#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
									open (SourceFile, "<$testfile");
									while($lineTF = <SourceFile>)												#reading parameter
									{
										$len = 0;
										chomp;
										$lineTF =~ s/^ +//;                               #clear head of line spacing
										#print $lineTF;
										if (substr($lineTF,0,8) eq "resistor")						####matching resistor########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,8);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											#print $parameter,'\n';
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,9) eq "capacitor")				####matching capacitor########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,9);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,8) eq "inductor")					####matching inductor########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,8);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,5) eq "diode")						####matching diode######
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,5);										#matching type
											#4.4# printf Tested "%-40s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													$tested-> write($rowT, 3, $HiLimit, $format_data);  ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													$tested-> write($rowT, 4, $LoLimit, $format_data);  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,5) eq "zener")						####matching zener##
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,5);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,6) eq "jumper")						####matching jumper########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,6);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
											$OP = 0; if ($lineTF =~ "op") {$OP = 1;}
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-52s", $Nominal;
													if($OP == 0)
													{$tested-> write($rowT, 3, $Nominal, $format_data);
													 $tested-> write($rowT, 5, $commentTP, $format_data);}  ## Excel ##
													if($OP == 1)
													{$tested-> write($rowT, 4, $Nominal, $format_STP);
													 $tested-> write($rowT, 5, "OP test", $format_STP);}	## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,4) eq "fuse")					  	####matching fuse########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,4);										#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  ## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-52s", $Nominal;
													$tested-> write($rowT, 3, $Nominal, $format_data);  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (eof)					  														####no parameter######
										{
										#4.4# printf Nulltested "%-30s", $testname; print Nulltested "No Test Parameter Found in TestFile."."\n";
										$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
										$untest-> write($rowU, 1, "No Test Parameter Found in TestFile.", $format_data);  ## Excel ##
										$rowU++;
											}
  								}
								}
							elsif (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname									#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}																		#ignore duplicated test name
  									$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n"; 
							$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
							$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							$rowU++;
							goto Next_Dev;}
							}
					close TesP;
				}
			################ multi-test(%)unpowered testable device ###########################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),1) eq "\%"
				and length(substr($lineTO,$StartBit_TO,index($lineTO,"\%")-$StartBit_TO)) == $len_device
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
					print "			Multi_Test(%)UNP  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
				 $testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
				 #print $testname,"\n";
					#print length($testname),"\n";
					open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching device name
								and substr($lineTP,0,1) ne "\!")							#matching not skipped test item
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}	
								#print $testfile,"\n";
									  if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
									open (SourceFile, "<$testfile");
									while($lineTF = <SourceFile>)							#reading parameter
									{
										$len = 0;
										chomp;
										$lineTF =~ s/^ +//;                               	#clear head of line spacing
										#print $lineTF;
										if (substr($lineTF,0,8) eq "resistor")				####matching resistor########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,8);						#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											#print $parameter,'\n';
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,9) eq "capacitor")				####matching capacitor########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,9);						#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,8) eq "inductor")				####matching inductor########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,8);						#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,5) eq "diode")					####matching diode######
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,5);						#matching type
											#4.4# printf Tested "%-40s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													$tested-> write($rowT, 3, $HiLimit, $format_data);  ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													$tested-> write($rowT, 4, $LoLimit, $format_data);  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,5) eq "zener")					####matching zener##
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,5);						#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-20s", $Nominal;
													$tested-> write($rowT, 2, $Nominal, $format_data);  ## Excel ##
													$HiLimit=shift(@array); #matching high limit
													#4.4# printf Tested "%-16s", $HiLimit;
													if($HiLimit < 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_data);}  ## Excel ##
													if($HiLimit >= 40)
													{$tested-> write($rowT, 3, $HiLimit, $format_STP);}   ## Excel ##
													$LoLimit=shift(@array); #matching low limit
													#4.4# printf Tested "%-16s", $LoLimit;
													if($LoLimit < 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_data);}  ## Excel ##
													if($LoLimit >= 40)
													{$tested-> write($rowT, 4, $LoLimit, $format_STP);}	  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,6) eq "jumper")					####matching jumper########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,6);						#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
											$OP = 0; if ($lineTF =~ "op") {$OP = 1;}
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-52s", $Nominal;
													if($OP == 0)
													{$tested-> write($rowT, 3, $Nominal, $format_data);
													 $tested-> write($rowT, 5, $commentTP, $format_data);}  ## Excel ##
													if($OP == 1)
													{$tested-> write($rowT, 4, $Nominal, $format_STP);
													 $tested-> write($rowT, 5, "OP test", $format_STP);}	## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (substr($lineTF,0,4) eq "fuse")					####matching fuse########
										{
											#4.4# printf Tested "%-30s", $testfile;
											$tested-> write($rowT, 0, $testfile, $format_item);  ## Excel ##
											$type = substr($lineTF,0,4);						#matching type
											#4.4# printf Tested "%-20s", $type;
											$tested-> write($rowT, 1, $type, $format_data);  	## Excel ##
											$parameter = substr($lineTF,index($lineTF,"\ ") + 1,length($lineTF));
											$parameter =~ s/\s//g;
											undef @array;
												while ($parameter =~ m/\,/g)
													{
														$paramt = substr($parameter,$len,pos($parameter) - $len - 1);
														push(@array,$paramt);
														$len = pos($parameter);
													}
													$Nominal=shift(@array); #matching nominal
													#4.4# printf Tested "%-52s", $Nominal;
													$tested-> write($rowT, 3, $Nominal, $format_data);  ## Excel ##
													#4.4# printf Tested "%-30s", $commentTP;  #TP comments
													$tested-> write($rowT, 5, $commentTP, $format_data);  ## Excel ##
													#4.8# print Tested "\n";
													$rowT++;
													last;
										}
										elsif (eof)					  						####no parameter######
										{
										#4.4# printf Nulltested "%-30s", $testname; print Nulltested "No Test Parameter Found in TestFile."."\n";
										$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
										$untest-> write($rowU, 1, "No Test Parameter Found in TestFile.", $format_data);  ## Excel ##
										$rowU++;
										}
  								}
								}
							elsif (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname	#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}					#ignore duplicated test name
  									$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n"; 
							$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
							$untest-> write($rowU, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							$rowU++;
							goto Next_Dev;}
							}
					close TesP;
				}
			################ multi-test(_)powered testable devices #############################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),1) eq "\_"
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO == -1
          and $skipTO == -1
          and $powered > -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					print "			Multi_Test(_)PWD  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# print PowerTest $lineTO,"\n";
					#4.6# $power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##

				 		$testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
						open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching device name
								and substr($lineTP,0,1) ne "\!")							#matching not skipped test item
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}	
									#print $testfile,"\n";
									$power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##
									$power-> write($rowP, 1, $lineTP, $format_data);  ## Excel ##
									$rowP++;
									  if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
  								}
							if (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname	#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}					#ignore duplicated test name
  									$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
								#	$power-> write($rowP, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								#	$power-> write($rowP, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_data);  ## Excel ##
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n"; 
						#	$power-> write($rowP, 0, $device, $format_data);  ## Excel ##
							$power-> write($rowP, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							goto Next_Dev;}
							}
						close TesP;

					#$rowP++;
				}
			################ multi-test(%)powered testable devices #############################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),1) eq "\%"
				and length(substr($lineTO,$StartBit_TO,index($lineTO,"\%")-$StartBit_TO)) == $len_device
          and $nullTO == -1
          and $skipTO == -1
          and $powered > -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1
          )
				{
					$foundTO = 1;
					print "			Multi_Test(%)PWD  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# print PowerTest $lineTO,"\n";
					#4.6# $power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##

				 		$testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
						open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching device name
								and substr($lineTP,0,1) ne "\!")							#matching not skipped test item
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}	
									#print $testfile,"\n";
									$power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##
									$power-> write($rowP, 1, $lineTP, $format_data);  ## Excel ##
									$rowP++;
									  if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
  								}
							if (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname	#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}					#ignore duplicated test name
  									$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
								#	$power-> write($rowP, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								#	$power-> write($rowP, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_data);  ## Excel ##
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n"; 
						#	$power-> write($rowP, 0, $device, $format_data);  ## Excel ##
							$power-> write($rowP, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							goto Next_Dev;}
							}
						close TesP;

					#$rowP++;
				}
			################ paralleled devices ##############################################################################################
			elsif(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO > -1
          and $comment1 > -1
          and $comment2 > -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Parallel_Test  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# printf Limited "%-30s", $device; print Limited substr($lineTO,index($lineTO,"\!"),length($lineTO)-index($lineTO,"\!")),"\n";
					$limited-> write($rowL, 0, $device, $format_data);  ## Excel ##
					$limited-> write($rowL, 1, substr($lineTO,index($lineTO,"\!")+1,length($lineTO)-index($lineTO,"\!")-1), $format_data);  ## Excel ##
					$rowL++;
				}
			################ paralleled multi-test (_)test name ##################################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),1) eq "\_"
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO > -1
          and $comment1 > -1
          and $comment2 > -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Parallel_Test(_)  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# printf Limited "%-30s", substr($lineTO,index($lineTO,$device), index($lineTO,"\;")- index($lineTO,$device)- 1); print Limited substr($lineTO,index($lineTO,"\!"),length($lineTO)-index($lineTO,"\!")),"\n";
					$limited-> write($rowL, 0, substr($lineTO,index($lineTO,$device), index($lineTO,"\;")- index($lineTO,$device)- 1), $format_data);  ## Excel ##
					$limited-> write($rowL, 1, substr($lineTO,index($lineTO,"\!")+1,length($lineTO)-index($lineTO,"\!")-1), $format_data);  ## Excel ##
					$rowL++;
				}
			################ paralleled multi-test (%)test name ##################################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),1) eq "\%"
				and length(substr($lineTO,$StartBit_TO,index($lineTO,"\%")-$StartBit_TO)) == $len_device
				and substr($lineTO,index($lineTO,$device)-1,1) eq "\""
          and $nullTO > -1
          and $comment1 > -1
          and $comment2 > -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Parallel_Test(%)  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# printf Limited "%-30s", substr($lineTO,index($lineTO,$device), index($lineTO,"\;")- index($lineTO,$device)- 1); print Limited substr($lineTO,index($lineTO,"\!"),length($lineTO)-index($lineTO,"\!")),"\n";
					$limited-> write($rowL, 0, substr($lineTO,index($lineTO,$device), index($lineTO,"\;")- index($lineTO,$device)- 1), $format_data);  ## Excel ##
					$limited-> write($rowL, 1, substr($lineTO,index($lineTO,"\!")+1,length($lineTO)-index($lineTO,"\!")-1), $format_data);  ## Excel ##
					$rowL++;
				}
			################ untestable devices ##############################################################################################
			elsif(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO > -1
          and $comment1 > -1
          and $comment2 == -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			NULL_Test  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# printf Nulltested "%-30s", $device; print Nulltested "set NullTest in TestOrder."; printf Nulltested "%-14s";
					$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
					$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_data);  ## Excel ##
				$UTline = "";
				open(ALL, "<analog/$device")||open(ALL, "<analog/1%$device") or $untest-> write($rowU, 2, "!TestFile not found.", $format_anno); #print Nulltested "!TestFile not found.\n";
				#open(ALL, "<analog/1%$device");
				while($line = <ALL>)
					{
					if (index($line,$device)>1){#4.4# print Nulltested $line; 
						#$untest-> write($rowU, 2, $line, $format_anno);
						$line = substr($line,1);
						$UTline = $line . $UTline;}
					elsif (eof){#4.4# print Nulltested "\n";
						#$untest-> write($rowU, 2, $UTline, $format_anno); 
						last;}
					}
					chomp($UTline);
					$untest-> write($rowU, 2, $UTline, $format_anno);
				$rowU++;
				close ALL;
				}
			################ multi-test(%)unpowered untestable device #################################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),1) eq "\%"
				and length(substr($lineTO,$StartBit_TO,index($lineTO,"\%")-$StartBit_TO)) == $len_device
          and $nullTO > -1
          || $comment1 > -1
          and $comment2 == -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Multi_UnTest(%)UNP  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
				$device1 = "";
				$device1 = substr($lineTO, index($lineTO,"\"")+ 1, rindex($lineTO,"\"")- index($lineTO,"\"")- 1);
				#4.4# printf Nulltested "%-30s", $device1; print Nulltested "set NullTest in TestOrder."; printf Nulltested "%-14s";
				$untest-> write($rowU, 0, $device1, $format_data);  ## Excel ##
				$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_data);  ## Excel ##
				#print $device1,"\n";
				$UTline = "";
				open(ALL, "<analog/$device1")||open(ALL, "<analog/1%$device1") or $untest-> write($rowU, 2, "!TestFile not found.", $format_anno); #print Nulltested "!TestFile not found.\n";
				#open(ALL, "<analog/1%$device");
				while($line = <ALL>)
					{
					#print $line,"\n";
					if (index($line,$device1)>1){#4.4# print Nulltested $line;
						#$untest-> write($rowU, 2, $line, $format_anno);
						$line = substr($line,1);
						$UTline = $line . $UTline;}
					elsif (index($line,"not accessible")>1){#4.4# print Nulltested $line; 
						#$untest-> write($rowU, 2, $line, $format_anno);
						$line = substr($line,1);
						$UTline = $line . $UTline;}
					elsif (index($line,"tested in file")>1){#4.4# print Nulltested $line; 
						#$untest-> write($rowU, 2, $line, $format_anno);
						$line = substr($line,1);
						$UTline = $line . $UTline;}
					elsif (eof){#4.4# print Nulltested "\n"; 
						last;}
					}
					chomp($UTline);
					$untest-> write($rowU, 2, $UTline, $format_anno);
				$rowU++;
				close ALL;
				}
			################ testable digital test ##########################################################################################
			elsif(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO == -1
          and $skipTO == -1
          and $powered == -1
          and $scan == -1
          and $digital > -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Digital_Test  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# print PowerTest $lineTO,"\n";
					#4.6# $power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##

				 		$testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
						open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching device name
								and substr($lineTP,0,1) ne "\!" and $lineTP =~ "test")		#matching not skipped test item
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len + 1);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}
									#print $lineTP,"\n";
									#print $testfile,"\n";
									$power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##
									$power-> write($rowP, 1, $lineTP, $format_data);  ## Excel ##
									#$rowP++;
									  if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
									
									$family = "";
  									open (SourceFile, "<$testfile");
									while($lineTF = <SourceFile>)							#reading family
									{
										chomp;
										$lineTF =~ s/^ +//;                               	#clear head of line spacing
										#print $lineTF;
										if (substr($lineTF,0,6) eq "family")
										{$family = $lineTF . $family;}
  										#{$power-> write($rowP, 2, $lineTF, $format_data);}  	## Excel ##
  									}
  									close SourceFile;
  									chomp($family);
  									$power-> write($rowP, 2, $family, $format_anno);
									$rowP++;
  								}
							if (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname	#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}					#ignore duplicated test name
  									$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
								#	$power-> write($rowP, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								#	$power-> write($rowP, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_data);  ## Excel ##
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n"; 
						#	$power-> write($rowP, 0, $device, $format_data);  ## Excel ##
							$power-> write($rowP, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							goto Next_Dev;}
							}
						close TesP;

					#$rowP++;
				}
			################ untestable digital test ##########################################################################################
			elsif(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO > -1
          || $skipTO > -1
          and $powered == -1
          and $scan == -1
          and $digital > -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Digital_UnTest  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# printf Nulltested "%-30s", $device; print Nulltested "set NullTest in TestOrder.              !Digital Test.\n";
					$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
					$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_data);  ## Excel ##
					$untest-> write($rowU, 2, "Digital Test.", $format_anno);  ## Excel ##
					$rowU++;
				}
			################ testable analog powered test ######################################################################################
			elsif(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO == -1
          and $skipTO == -1
          and $powered > -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			ANA_PWD_Test  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# print PowerTest $lineTO,"\n";
					#4.6# $power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##

				 		$testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
						open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching device name
								and substr($lineTP,0,1) ne "\!")							#matching not skipped test item
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}	
									#print $testfile,"\n";
									$power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##
									$power-> write($rowP, 1, $lineTP, $format_data);  ## Excel ##
									$rowP++;
									  if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
  								}
							if (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname	#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}					#ignore duplicated test name
  									$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
								#	$power-> write($rowP, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								#	$power-> write($rowP, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_data);  ## Excel ##
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n"; 
						#	$power-> write($rowP, 0, $device, $format_data);  ## Excel ##
							$power-> write($rowP, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							goto Next_Dev;}
							}
						close TesP;

					#$rowP++;
				}
			################ untestable analog powered test ######################################################################################
			elsif(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO > -1
          || $skipTO > -1
          and $powered > -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			ANA_PWD_UnTest  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# printf Nulltested "%-30s", $device; print Nulltested "set NullTest in TestOrder.              !Analog Powered Test.\n";
					$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
					$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_data);  ## Excel ##
					#4.9# $untest-> write($rowU, 2, "Analog Powered Test.", $format_anno);  ## Excel ##
				$UTline = "";
				open(ALL, "<analog/$device")||open(ALL, "<analog/1%$device") or $untest-> write($rowU, 2, "!TestFile not found.", $format_anno); #print Nulltested "!TestFile not found.\n";
				#open(ALL, "<analog/1%$device");
				while($line = <ALL>)
					{
					if (index($line,"not accessible")>1){#4.4# print Nulltested $line; 
						#$untest-> write($rowU, 2, $line, $format_anno);
						$line = substr($line,6);
						$UTline = $line . $UTline;}
					elsif (eof){#4.4# print Nulltested "\n";
						#$untest-> write($rowU, 2, $UTline, $format_anno); 
						last;}
					}
					chomp($UTline);
					$untest-> write($rowU, 2, $UTline, $format_anno);
					$rowU++;
					close ALL;
				}
			################ multi-test(%)powered untestable device ####################################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),1) eq "\%"
				and length(substr($lineTO,$StartBit_TO,index($lineTO,"\%")-$StartBit_TO)) == $len_device
          and $nullTO > -1
          || $skipTO > -1
          and $powered > -1
          and $scan == -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Multi_UnTest(%)PWD  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
				$device1 = "";
				$device1 = substr($lineTO, index($lineTO,"\"")+ 1, rindex($lineTO,"\"")- index($lineTO,"\"")- 1);
				#4.4# printf Nulltested "%-30s", $device1; print Nulltested "set NullTest in TestOrder."; printf Nulltested "%-14s";
				$untest-> write($rowU, 0, $device1, $format_data);  ## Excel ##
					$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_data);  ## Excel ##
				#print $device1,"\n";
				$UTline = "";
				open(ALL, "<analog/$device1")||open(ALL, "<analog/1%$device1") or $untest-> write($rowU, 2, "!TestFile not found.", $format_anno); #print Nulltested "!TestFile not found.\n";
				#open(ALL, "<analog/1%$device");
				while($line = <ALL>)
					{
					#print $line,"\n";
					if (index($line,$device1)>1){#4.4# print Nulltested $line;
						#$untest-> write($rowU, 2, $line, $format_anno);
						$line = substr($line,6);
						$UTline = $line . $UTline;}
					elsif (index($line,"not accessible")>1){#4.4# print Nulltested $line; 
						#$untest-> write($rowU, 2, $line, $format_anno);
						$line = substr($line,6);
						$UTline = $line . $UTline;}
					elsif (index($line,"tested in file")>1){#4.4# print Nulltested $line; 
						#$untest-> write($rowU, 2, $line, $format_anno);
						$line = substr($line,6);
						$UTline = $line . $UTline;}
					elsif (eof){#4.4# print Nulltested "\n"; 
						last;}
					}
					chomp($UTline);
					$untest-> write($rowU, 2, $UTline, $format_anno);
					$rowU++;
				close ALL;
				}
			################ testable mixed device ##########################################################################################
			elsif(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO == -1
          and $skipTO == -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed > -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Mixed_Test  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# print PowerTest $lineTO,"\n";
					#4.6# $power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##

				 		$testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
						open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching device name
								and substr($lineTP,0,1) ne "\!")							#matching not skipped test item
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}	
									#print $testfile,"\n";
									$power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##
									$power-> write($rowP, 1, $lineTP, $format_data);  ## Excel ##
									$rowP++;
									  if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
  								}
							if (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname	#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}					#ignore duplicated test name
  									$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
								#	$power-> write($rowP, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								#	$power-> write($rowP, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_data);  ## Excel ##
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n"; 
						#	$power-> write($rowP, 0, $device, $format_data);  ## Excel ##
							$power-> write($rowP, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							goto Next_Dev;}
							}
						close TesP;

					#$rowP++;
				}
			################ untestable mixed device ##########################################################################################
			elsif(substr($lineTO,$StartBit_TO,$StopBit_TO) eq $device
				and length(substr($lineTO,$StartBit_TO,$StopBit_TO)) == $len_device
          and $nullTO > -1
          || $skipTO > -1
          and $powered == -1
          and $scan == -1
          and $digital == -1
          and $mixed > -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Mixed_UnTest  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# printf Nulltested "%-30s", $device; print Nulltested "set NullTest in TestOrder.\n";
					$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
					$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_data);  ## Excel ##
					$rowU++;
				}
			################ testable Bscan device #######################################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),8) eq "_connect"
				and length(substr($lineTO,$StartBit_TO,index($lineTO,"_connect")-$StartBit_TO)) == $len_device
          and $nullTO == -1
          and $skipTO == -1
          and $powered == -1
          and $scan > -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Bscan_Test  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# print PowerTest $lineTO,"\n";
					#4.6# $power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##

				 		$testname = substr($lineTO,index($lineTO,"\"") + 1, rindex($lineTO,"\"") - index($lineTO,"\"") - 1);
						open (TesP, "<testplan");												#### check testplan ####
						while($lineTP = <TesP>)
							{
							chomp($lineTP);
							$lineTP =~ s/^ +//;												#clear head of line spacing
							undef @array;
							while ($lineTP =~ m/\"/g)
								{
									$leng = pos($lineTP);
									#print $leng,"\n";
									push(@array,$leng);
								}
								if (index($lineTP,"\/")>0)									#matching single version beginning bit
									{
									$StartBit_Dummy = shift(@array)+1;
									$StartBit_TP = index($lineTP,"\/")	+ 1;
									$StopBit = shift(@array);
									$StopBit_TP = $StopBit - $StartBit_TP - 1; 				#matching stop bit
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP + $StartBit_Dummy;
									#print $device_len,"\n";
									}
								elsif(index($lineTP,"\/")<0)								#matching multi-version beginning bit
									{
									$StartBit_TP = shift(@array);
									$Stop_TP = shift(@array); 								#matching stop bit
									$StopBit_TP = $Stop_TP - $StartBit_TP - 1;
									#print $StartBit_TP,"\n";
									#print $StopBit_TP,"\n";
									$device_len = $StopBit_TP;
									#print $device_len,"\n";
									}

									#print substr($lineTP,$StartBit_TP,$StopBit_TP),"\n";
							if(substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname        #matching device name
								and substr($lineTP,0,1) ne "\!")							#matching not skipped test item
								{
									$foundTP = 1;
								#print "  ",$lineTP,"\n";
								if (index($lineTP,"on boards")== -1)
									{
									$testfile = substr($lineTP,index($lineTP,"\"") + 1, $device_len + 1);
									}
								if (index($lineTP,"on boards")> -1)
									{
									$testfile = "analog/1%".$testname;
									}
									#print $lineTP,"\n";
									#print $testfile,"\n";
									$power-> write($rowP, 0, $lineTO, $format_data);  ## Excel ##
									$power-> write($rowP, 1, $lineTP, $format_data);  ## Excel ##
									#$rowP++;
									  if ($testfile eq $testfile_last){last;}				#ignore duplicated test name
									  $commentTP = "";
									  if (index($lineTP,"\!")> -1) {$commentTP = substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"));}   #TP comments
  									$testfile_last = $testfile;
									
									$family = "";
  									open (SourceFile, "<$testfile");
									while($lineTF = <SourceFile>)							#reading family
									{
										chomp;
										$lineTF =~ s/^ +//;                               	#clear head of line spacing
										#print $lineTF;
										if (substr($lineTF,0,6) eq "family")
										{$family = $lineTF . $family;}
  										#{$power-> write($rowP, 2, $lineTF, $format_data);}  	## Excel ##
  									}
  									close SourceFile;
  									chomp($family);
  									$power-> write($rowP, 2, $family, $format_anno);
									$rowP++;
  								}
							if (substr($lineTP,$StartBit_TP,$StopBit_TP) eq $testname	#matching skipped test in TP
								and substr($lineTP,0,1) eq "\!")
								{
									$foundTP = 1;
									#print Nulltested $lineTP, "\n";
									if ($testname eq $testname_last){last;}					#ignore duplicated test name
  									$testname_last = $testname;
									#4.4# printf Nulltested "%-30s", $testname; print Nulltested "been skipped in TestPlan."; printf Nulltested "%-15s"; print Nulltested substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!"))."\n";
								#	$power-> write($rowP, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 0, $testname, $format_data);  ## Excel ##
									$untest-> write($rowU, 1, "been skipped in TestPlan.", $format_STP);  ## Excel ##
									$untest-> write($rowU, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_anno);  ## Excel ##
									$rowU++;
								#	$power-> write($rowP, 2, substr($lineTP,rindex($lineTP,"\!"),length($lineTP)- rindex($lineTP,"\!")), $format_data);  ## Excel ##
								}
							elsif (eof and $foundTP == 0){#4.4# printf Nulltested "%-30s", $device; print Nulltested "NO test item found in TestPlan.\n"; 
						#	$power-> write($rowP, 0, $device, $format_data);  ## Excel ##
							$power-> write($rowP, 1, "NO test item found in TestPlan.", $format_STP);  ## Excel ##
							goto Next_Dev;}
							}
						close TesP;

					#$rowP++;
				}
			################ untestable Bscan device #####################################################################################
			elsif(substr($lineTO,index($lineTO,$device) + length($device),8) eq "_connect"
				and length(substr($lineTO,$StartBit_TO,index($lineTO,"_connect")-$StartBit_TO)) == $len_device
          and $nullTO > -1
          || $skipTO > -1
          and $powered == -1
          and $scan > -1
          and $digital == -1
          and $mixed == -1
          and $ver == -1)
				{
					$foundTO = 1;
					print "			Bscan_UnTest  ", substr($lineTO,index($lineTO,$device)-1, length($lineTO)-index($lineTO,$device)+1),"\n";   #, $lineTO,"\n";
					#4.4# printf Nulltested "%-30s", substr($lineTO,index($lineTO,$device),index($lineTO,"\;")-index($lineTO,$device)-1); print Nulltested "set NullTest in TestOrder.              !Bscan Test.\n";
					$untest-> write($rowU, 0, substr($lineTO,index($lineTO,$device),index($lineTO,"\;")-index($lineTO,$device)-1), $format_data);  ## Excel ##
					$untest-> write($rowU, 1, "been set NullTest in TestOrder.", $format_data);  ## Excel ##
					$untest-> write($rowU, 2, "Bscan Test.", $format_anno);  ## Excel ##
					$rowU++;
				}
      ################ reservation #######################################################################################################
      elsif (eof and $foundTO == 0)
      	{
      		print "			NO Test Item Found\n"; #4.4# printf Nulltested "%-30s", $device; 
      		#4.4# print Nulltested "NO test item found in TestOrder,    !Check TJ/SP testing.\n";
			$untest-> write($rowU, 0, $device, $format_data);  ## Excel ##
			$untest-> write($rowU, 1, "NO test item found in TestOrder.", $format_data);  ## Excel ##
			$untest-> write($rowU, 2, "Check TJ/SP testing.", $format_anno);  ## Excel ##
			$rowU++;
      		goto Next_Dev;
      		}
      #############################################################################################################################
			}
Next_Dev:
}


############################### shorts threshold statistic ########################################################################

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
				$thres =~ s/\s//g;                     #clear all spacing
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

########################################################@##########################################################################


#4.4# close Tested;
#4.4# close Nulltested;
#4.4# close Limited;
#4.4# close PowerTest;
$bom_coverage_report->close();

print  "\n  >>> Completed ...\n";

END_Prog:

print "\n";
system 'pause';
exit;

