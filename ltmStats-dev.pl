#! /usr/bin/env perl

#####################################################
# Collects numerous statistics from a target BIP-IP while under test
# writes the output to an excel spreadsheet
#
# Copyright, F5 Networks, 2009-2015
# Written by: Jesse Driskill, Sr. ITC Systems Engineer
#####################################################

## Required libraries
#
use warnings;
use strict;

use Config;
use Getopt::Std;
use Net::SNMP     qw(:snmp);
use Time::HiRes   qw(tv_interval gettimeofday ualarm usleep time);
use Data::Dumper;
use JSON;
use Clone         qw/clone/;       
use List::Util    qw/sum/;
use Excel::Writer::XLSX;

$!++;  # No buffer for STDOUT
$SIG{'INT'} = \&exit_now;  # handle ^C nicely

## retrieve and process CLI paramters
#
our (%opts, $XLSXOUT, $JSONOUT, $BYPASS, $DEBUG, $VERBOSE, $CUSTOMER);
our ($TESTNAME, $COMMENTS, $PAUSE);
getopts('d:m:l:o:j:C:T:m:c:s:i:p:BDvh', \%opts);

# print usage and exit
&usage(0) if $opts{'h'};

if (!$opts{'d'}) {
  warn("Must provide a hostname or IP address to query\n");
  &usage(1);
}

my $host      = $opts{'d'};                     # snmp host to poll
my $secondary = $opts{'m'};                     # monitoring host
my $testLen   = $opts{'l'} || 130;              # total duration of test in seconds
my $xlsxName  = $opts{'o'} || '/dev/null';      # xlsx output file name
my $jsonName  = $opts{'j'} || '/dev/null';      # json output file name
my $snmpVer   = $opts{'s'} || 'v2c';            # snmp version
my $comm      = $opts{'c'} || 'public';         # community string    
my $cycleTime = $opts{'i'} || 10;               # polling interval
my $pause     = $opts{'p'} || 0;                # pause time
my $customer  = $opts{'C'} || 'not provided';   # Customer name
my $testname  = $opts{'T'} || 'not provided';   # Test name
my $comments  = $opts{'m'} || 'not provided';   # Test comments/description


# The signal handler will throw an error if the output files ($XLSXOUT and $JSONOUT)
# aren't defined. Cosmetic, but irritating.
$DEBUG      = ($opts{'D'} ? 1 : 0);
$XLSXOUT    = ($opts{'o'} ? 1 : 0);
$JSONOUT    = ($opts{'j'} ? 1 : 0);
$VERBOSE    = ($opts{'v'} ? 1 : 0);
$BYPASS     = ($opts{'B'} ? 1 : 0);
$CUSTOMER   = ($opts{'C'} ? 1 : 0);
$TESTNAME   = ($opts{'T'} ? 1 : 0);
$COMMENTS   = ($opts{'m'} ? 1 : 0);


if ($DEBUG) {
  print Dumper(\%opts);
  #($_ = <<EOF) =~ s/^\s+//gm;
  print "
  XLSXOUT:  $XLSXOUT
  BYPASS:   $BYPASS
  DEBUG:    $DEBUG
  VERBOSE:  $VERBOSE
  CUSTOMER: $CUSTOMER
  TESTNAME: $TESTNAME
  COMMENTS: $COMMENTS"
#EOF
}

# additional constants
use constant MB  => 1024*1024;
use constant GB  => 1024*1024*1024;
my $usCycleTime  = $cycleTime * 1_000_000;

##
## Initialization and environment check
##

my (@dataList, @errorList, @staticList, @rowData, %formats);
my ($clientCurConns, $clientTotConns, $serverCurConns, $serverTotConns);
my ($cpuUsed, $cpuTicks, $cpuUtil, $cpuPercent, $tmmUtil, $tmmPercent);
my ($memUsed, $hMem, $dataVals, $errorVals, $col);
my ($workbook, $summary, $raw_data, $chtdata, $charts);
my ($cBytesIn, $cBytesOut, $sBytesIn, $sBytesOut, $tBytesIn, $tBytesOut);
my ($cPktsIn, $cPktsOut, $sPktsIn, $sPktsOut);
my ($cNewConns, $sNewConns, $ccPktsIn, $ccPktsOut, $cBitsIn, $cBitsOut)   = (0, 0, 0, 0, 0, 0);
my ($row, $sBitsIn, $sBitsOut, $tBitsIn, $tBitsOut, $httpReq)             = (0, 0, 0, 0, 0, 0);
my ($slept, $sleepTime, $pollTime, $runTime, $lastLoopEnd, $loopTime)     = (0, 0, 0, 0, 0, 0);

my ($old, $cur, $out, $xlsData, $test_meta) = ({}, {}, {}, {}, {});
my %pollTimer = ();      # contains event timestamps

$test_meta->{customer}  = "$customer";
$test_meta->{test_name} = "$testname";
$test_meta->{comments}  = "$comments";

# Build the oid lists and varbind arrays
my %staticOids  = &get_static_oids();
my %dataOids    = &get_f5_oids();
my %errorOids   = &get_err_oids();

my @dutInfoHdrs = qw(Host Platform Version Build Memory CPUs Blades);
my @chtDataHdrs = ('RunTime', 'SysCPU', 'TmmCPU', 'Memory', 'Client Mbs In', 
                   'Client Mbs Out', 'Server Mbs In', 'Server Mbs Out','Client CurConns',
                   'Server CurConns', 'Client Conns/Sec', 'Server Conns/Sec',
                   'HTTP Requests/Sec', 'Total CurConns', 
                  );
my @summaryHdrs = ('RunTime', 'LoopTime', 'SysCPU', 'TmmCPU', 'Memory', 'Client bitsIn/s', 
                   'Client bitsOut/s', 'Server bitsIn/s', 'Server bitsOut/s', 
                   'Client pktsIn/s', 'Client pktsOut/s', 'Server pktsIn/s', 'Server pktsOut/s',
                   'Client Conn/s', 'Server Conn/s', 'HTTP Requests/Sec',
                  );
my @rawdataHdrs = ('RunTime', 'SysCPU', 'TmmCPU', 'Memory', 'Client bytesIn', 'Client bytesOut', 
                   'Client pktsIn', 'Client pktsOut', 'Server btyesIn', 'Server bytesOut', 
                   'Server pktsIn', 'Server pktsOut', 'Client curConns', 'Client totConns', 
                   'Server curConns', 'Server totConns', 'HTTP Requests',
                   'Clt PVA Bytes In', 'Clt PVA Bytes Out', 'Clt TMM Bytes In', 'Clt TMM Bytes Out',
                   'Svr PVA Bytes In', 'Svr PVA Bytes Out', 'Svr TMM Bytes In', 'Svr TMM Bytes Out',
                  );

while (my ($key, $value) = each(%staticOids)) { push(@staticList, $value); }
while (my ($key, $value) = each(%dataOids))   { push(@dataList, $value); }
while (my ($key, $value) = each(%errorOids))  { push(@errorList, $value); }

my ($session, $error) = Net::SNMP->session(
  -hostname     => $host,
  -community    => $comm,
  -version      => $snmpVer,
  -maxmsgsize   => 8192,
  -nonblocking  => 0,
);
die($error."\n") if ($error);

# determine if logging is required and create the output files
if ($XLSXOUT) {
  $DEBUG && print "Creating workbook ($xlsxName)\n";
  ($workbook, $raw_data, $summary, $chtdata, $charts, %formats) = 
      &mk_perf_xls($xlsxName, \@rawdataHdrs, \@summaryHdrs, \@chtDataHdrs, \@dutInfoHdrs);
}

# print out some information about the DUT being polled
my $result = $session->get_request( -varbindlist  => \@staticList);
print "Platform:    $result->{$staticOids{platform}}\n";
print "Memory:      $result->{$staticOids{totalMemory}} (".$result->{$staticOids{totalMemory}} / MB." MB)\n";
print "# of CPUs:   $result->{$staticOids{cpuCount}}\n";
print "# of blades: $result->{$staticOids{bladeCount}}\n";
print "LTM Version: $result->{$staticOids{ltmVersion}}\n";
print "LTM Build:   $result->{$staticOids{ltmBuild}}\n";

# If a real xlsx is being written to, record DUT vital info on the first sheet
if ($xlsxName !~ '/dev/null') {
  #while (my ($k, $v) = each(%staticOids)) {
  #  print $k.": ".$result->{$v}."\n";
  #}
  $charts->write("A2", $result->{$staticOids{hostName}},    $formats{text});
  $charts->write("B2", $result->{$staticOids{platform}},    $formats{text});
  $charts->write("C2", $result->{$staticOids{ltmVersion}},  $formats{text});
  $charts->write("D2", $result->{$staticOids{ltmBuild}},    $formats{text});
  $charts->write("E2", $result->{$staticOids{totalMemory}}, $formats{decimal0});
  $charts->write("F2", $result->{$staticOids{cpuCount}},    $formats{text});
  $charts->write("G2", $result->{$staticOids{bladeCount}},  $formats{text});
}


$test_meta->{host_name}     = $result->{$staticOids{hostName}};
$test_meta->{platform}      = $result->{$staticOids{platform}};
$test_meta->{cpu_count}     = $result->{$staticOids{cpuCount}};
$test_meta->{blade_count}   = $result->{$staticOids{bladeCount}};
$test_meta->{memory}        = $result->{$staticOids{totalMemory}};
$test_meta->{ltm_version}   = $result->{$staticOids{ltmVersion}};
$test_meta->{ltm_build}     = $result->{$staticOids{ltmBuild}};

my %json_buffer = ( 'metadata' => $test_meta,
                    'perfdata' => [],
                  );

##
## Begin Main
##

# loop until start-of-test is detected
if ($opts{'m'}) {
  my ($watchhost, $error) = Net::SNMP->session(
    -hostname     => $secondary,
    -community    => $comm,
    -version      => $snmpVer,
    -maxmsgsize   => 8192,
    -nonblocking  => 0,
  );
  die($error."\n") if ($error);
  &detect_test($watchhost, \%dataOids) unless $BYPASS;
}
else {
  &detect_test($session, \%dataOids) unless $BYPASS;
}

if ($pause) {
#if ($pause && !$BYPASS) {
  print "Pausing for ".$pause." seconds while for ramp-up\n";
  sleep($pause);
}

# start active polling
$pollTimer{testStart} = [gettimeofday];

do {
  my $out = {};
  $pollTimer{activeStart} = [gettimeofday];


  # get snmp stats from DUT
  $dataVals = $session->get_request( -varbindlist  => \@dataList);
  die($session->error."\n") if (!defined($dataVals));

  $pollTimer{queryTime}     = tv_interval($pollTimer{activeStart});
  $pollTimer{lastPollTime}  = $pollTimer{pollTime};
  $pollTimer{pollTime}      = [gettimeofday];


  # Confirm at least one full iteration has completed
  if ($pollTimer{lastLoopEnd}) {
    $loopTime = tv_interval($pollTimer{lastPollTime});
    $runTime  = tv_interval($pollTimer{testStart});
  } else {
    $loopTime = 0;
  }

  # Before any real processing, remove any non-numeric values (i.e. 'noSuchInstance')
  for my $n (keys(%dataOids)) {
    if ($dataVals->{$dataOids{$n}} =~ /\D+/) {
      $xlsData->{$n} = 1;
    }
    else {
      $xlsData->{$n} = $dataVals->{$dataOids{$n}};
    }
  }

  # CPU and TMM time used
  $cur->{cpuTotalTicks}   = sum(@$xlsData{qw/ssCpuRawUser ssCpuRawNice ssCpuRawSystem ssCpuRawIdle/});
  $cur->{cpuIdleTicks}    = $xlsData->{ssCpuRawIdle};
  $cur->{tmmTotal}        = $xlsData->{tmmTotalCycles};
  $cur->{tmmIdle}         = sum(@$xlsData{qw/tmmIdleCycles tmmSleepCycles/});
  $cur->{memUsed}         = $xlsData->{tmmTotalMemoryUsed};
  $cur->{clientCurConns}  = $xlsData->{tmmClientCurConns};
  $cur->{serverCurConns}  = $xlsData->{tmmServerCurConns};
  $cur->{clientTotConns}  = $xlsData->{tmmClientTotConns};
  $cur->{serverTotConns}  = $xlsData->{tmmServerTotConns};
  $cur->{cPvaBytesIn}     = $xlsData->{pvaClientBytesIn};   # Client PVA bytes in
  $cur->{cPvaBytesOut}    = $xlsData->{pvaClientBytesOut};  # Client PVA bytes out
  $cur->{cTmmBytesIn}     = $xlsData->{tmmClientBytesIn};   # Client TMM bytes in
  $cur->{cTmmBytesOut}    = $xlsData->{tmmClientBytesOut};  # Client TMM bytes out
  $cur->{sPvaBytesIn}     = $xlsData->{pvaServerBytesIn};   # Server PVA bytes in
  $cur->{sPvaBytesOut}    = $xlsData->{pvaServerBytesOut};  # Server PVA bytes out
  $cur->{sTmmBytesIn}     = $xlsData->{tmmServerBytesIn};   # Server TMM bytes in
  $cur->{sTmmBytesOut}    = $xlsData->{tmmServerBytesOut};  # Server TMM bytes out
  $cur->{cBytesIn}        = $xlsData->{tmmClientBytesIn}  + $xlsData->{pvaClientBytesIn};
  $cur->{cBytesOut}       = $xlsData->{tmmClientBytesOut} + $xlsData->{pvaClientBytesOut};
  $cur->{sBytesIn}        = $xlsData->{tmmServerBytesIn}  + $xlsData->{pvaServerBytesIn};
  $cur->{sBytesOut}       = $xlsData->{tmmServerBytesOut} + $xlsData->{pvaServerBytesOut};
  $cur->{cPktsIn}         = $xlsData->{tmmClientPktsIn}   + $xlsData->{pvaClientPktsIn};
  $cur->{cPktsOut}        = $xlsData->{tmmClientPktsOut}  + $xlsData->{pvaClientPktsOut};
  $cur->{sPktsIn}         = $xlsData->{tmmServerPktsIn}   + $xlsData->{pvaServerPktsIn};
  $cur->{sPktsOut}        = $xlsData->{tmmServerPktsOut}  + $xlsData->{pvaServerPktsOut};
  $cur->{totHttpReq}      = $xlsData->{sysStatHttpRequests};

  if ($runTime) { 
    $out->{runTime}       = $runTime;
    $out->{httpReq}       = sprintf("%.0f", delta("totHttpReq") / $loopTime);
    $out->{cNewConns}     = sprintf("%.0f", delta("clientTotConns") / $loopTime);
    $out->{sNewConns}     = sprintf("%.0f", delta("serverTotConns") / $loopTime);
    $out->{cCurConns}     = sprintf("%.0f", $cur->{clientCurConns});
    $out->{sCurConns}     = sprintf("%.0f", $cur->{serverCurConns});
    $out->{cBitsIn}       = sprintf("%.0f", bytes_to_Mbits(delta("cBytesIn"))  / $loopTime);
    $out->{cBitsOut}      = sprintf("%.0f", bytes_to_Mbits(delta("cBytesOut")) / $loopTime);
    $out->{sBitsIn}       = sprintf("%.0f", bytes_to_Mbits(delta("sBytesIn"))  / $loopTime);
    $out->{sBitsOut}      = sprintf("%.0f", bytes_to_Mbits(delta("sBytesOut")) / $loopTime);
    $out->{cPktsIn}       = sprintf("%.0f", delta("cPktsIn")  / $loopTime);
    $out->{cPktsOut}      = sprintf("%.0f", delta("cPktsOut") / $loopTime);
    $out->{sPktsIn}       = sprintf("%.0f", delta("sPktsIn")  / $loopTime);
    $out->{sPktsOut}      = sprintf("%.0f", delta("sPktsOut") / $loopTime);
    $out->{memUsed}       = sprintf("%.0f", bytes_to_MB($cur->{memUsed}));
    $out->{cpuUtil}       = sprintf("%.2f", cpu_util(delta("cpuTotalTicks"), delta("cpuIdleTicks")));
    $out->{tmmUtil}       = sprintf("%.2f", cpu_util(delta("tmmTotal"), delta("tmmIdle")));


    if ($VERBOSE) {
    # This 'format' displays the standard data
#@|||||||| @>>>>  @>>>>   @|||||||  @|||||  @|||||   @||||||  @|||||||||  @|||||||||  @||||||  @||||||| @|||||||| @||||||||
    format STDOUT_TOP =
@>>>>>>>> @>>>>  @>>>>   @>>>>>>>  @>>>>>>  @>>>>>> @>>>>>>  @>>>>>>>>>  @>>>>>>>>> @>>>>>>  @>>>>>> @>>>>>>>> @>>>>>>>>
"Time", "CPU", "TMM", "Mem (MB)", "C-CPS", "S-CPS", "HTTP", "Client CC", "Server CC", "In/Mbs", "Out/Mbs", "cPPS In", "cPPS Out"
.

    format =
@####.###  @#.##  @#.##  @#####  @######  @######  @######  @#########  @#########  @######  @###### @######## @########
@$out{qw/runTime cpuUtil tmmUtil memUsed cNewConns sNewConns httpReq cCurConns sCurConns cBitsIn cBitsOut cPktsIn cPktsOut/}
.
      write;
      #printf("%.3d %.2d %.2d %6d %8d %8d %8d %9d %9d %6d %6d %9d %9d\n", 
      #    @$out{qw/runTime cpuUtil tmmUtil memUsed cNewConns sNewConns httpReq cCurConns sCurConns cBitsIn cBitsOut cPktsIn cPktsOut/})
    }
#    else {
#    # This 'format' displays the standard data
#    format STDOUT_TOP =
#@|||||||| @>>>>  @>>>>   @|||||||  @|||||  @|||||   @||||||  @|||||||||  @|||||||||  @||||||  @|||||||
#"Time", "CPU", "TMM", "Mem (MB)", "C-CPS", "S-CPS", "HTTPReq", "Client CC", "Server CC", "In/Mbs", "Out/Mbs"
#.
#
#    format =
#@####.###  @#.## @##.##  @#####  @######  @######  @######  @#########  @#########  @######  @######
#@$out{qw/runTime cpuUtil tmmUtil memUsed cNewConns sNewConns httpReq cCurConns sCurConns cBitsIn cBitsOut/}
#.
#      write;
#    }

    # If requested, write the output file.
    if ($XLSXOUT) {
      $row++;
      $raw_data->write($row, 0, $out->{runTime}, $formats{decimal4});
      $raw_data->write($row, 1, $out->{cpuUtil}, $formats{decimal2});
      $raw_data->write($row, 2, $out->{tmmUtil}, $formats{decimal2});

      $raw_data->write( $row, 
                        3,
                        [$cur->{memUsed},
                        sprintf("%.0f", $cur->{cBytesIn}),
                        sprintf("%.0f", $cur->{cBytesOut}),
                        sprintf("%.0f", $cur->{cPktsIn}),
                        sprintf("%.0f", $cur->{cPktsOut}),
                        sprintf("%.0f", $cur->{sBytesIn}),
                        sprintf("%.0f", $cur->{sBytesOut}),
                        sprintf("%.0f", $cur->{sPktsIn}),
                        sprintf("%.0f", $cur->{sPktsOut}),
                        sprintf("%.0f", $cur->{clientCurConns}),
                        sprintf("%.0f", $cur->{clientTotConns}),
                        sprintf("%.0f", $cur->{serverCurConns}),
                        sprintf("%.0f", $cur->{serverTotConns}),
                        sprintf("%.0f", $cur->{totHttpReq}),
                        sprintf("%.0f", $cur->{cPvaBytesIn}),
                        sprintf("%.0f", $cur->{cPvaBytesOut}),
                        sprintf("%.0f", $cur->{cTmmBytesIn}),
                        sprintf("%.0f", $cur->{cTmmBytesOut}),
                        sprintf("%.0f", $cur->{sPvaBytesIn}),
                        sprintf("%.0f", $cur->{sPvaBytesOut}),
                        sprintf("%.0f", $cur->{sTmmBytesIn}),
                        sprintf("%.0f", $cur->{sTmmBytesOut})],
                       $formats{'standard'});
    }

    # Save data in json_buffer in case that output has been requested
    # Make sure we 'numify' the data-points before writing them out
    foreach my $k (keys %$out) { $out->{$k} += 0; }
    push(@{$json_buffer{perfdata}}, $out);
  }

  # update 'old' data with the current values to calculate delta next cycle
  $old = clone($cur);

  # Calculate how much time this iteration has required to determine how
  # long we should sleep before beginning the next cycle
  $pollTimer{iterationTime} = tv_interval($pollTimer{activeStart});
  $sleepTime = $cycleTime - $pollTimer{iterationTime};

  my $wakeTime = Time::HiRes::time() + $sleepTime;

  $DEBUG && printf("Query: %.6f, iteration: %.6f, sleep: %.6f, loop: %.6f\n",
              $pollTimer{queryTime}, $pollTimer{iterationTime}, $sleepTime, $loopTime);

  $pollTimer{lastLoopEnd} = [gettimeofday];
  while (Time::HiRes::time() < $wakeTime) {
    #print "Now: ".Time::HiRes::time()."  Waketime: $wakeTime\n";
    Time::HiRes::usleep(5);
    #next;
  }
} while ($runTime < $testLen);

# polling is now complete, time to write the output files (if requested)
if ( $JSONOUT) {
  print "Writing JSON output file: $jsonName\n";
  &json_fwrite();
}

if ($XLSXOUT) {
  print "Writing XLSX output file: $xlsxName\n";
  #&write_summary($summary, \%formats, $row);
  &write_chartData($chtdata, \%formats, $row);
  &mk_charts($workbook, $charts, $row);

  # close the workbook; required for the workbook to be usable.
  &close_xls($workbook);
}


##
## Subs
##

# utility subs
sub delta           { return (!exists $old->{$_[0]} ? 0 : $cur->{$_[0]} - $old->{$_[0]}); }
sub delta2          {
  printf("Old: %.0f   Cur: %.0f   Diff: %.0f\n", $old->{$_[0]}, $cur->{$_[0]}, $cur->{$_[0]} - $old->{$_[0]});
  return ((!exists $old->{$_[0]}) ? 0 : ($cur->{$_[0]} - $old->{$_[0]}));
};
sub bytes_to_Mbits  { return sprintf("%.0f", ($_[0] * 8) / 1_000_000) };
sub bytes_to_MB     { return sprintf("%d", $_[0] / (1024 * 1024)) };
sub cpu_util        { return ($_[0] == 0) ? 0 : (($_[0] - $_[1]) / $_[0]) * 100 };

# delay the start of the script until the test is detected through pkts/sec
sub detect_test() {
  my $snmp = shift;
  my $oids = shift;
  my $pkts = 1000;

  print "\nWaiting for test to begin...\n";

  while (1) {
    my $r1 = $snmp->get_request($$oids{tmmClientPktsIn});
    sleep(4);
    my $r2 = $snmp->get_request($$oids{tmmClientPktsIn});

    my $delta = $r2->{$$oids{tmmClientPktsIn}}- 
                $r1->{$$oids{tmmClientPktsIn}};
  
    if ($delta > $pkts) {
      print "Start of test detected...\n\n";
      return;
    }
  }
}

# write the formulas in the summary sheet. 
# IN:   $row  - number of data rows in 'raw_data' worksheet
# OUT:  nothing
sub write_summary() {
  my $worksheet = shift;
  my $formats   = shift;
  my $numRows   = shift;
  my ($row0, $col, $row1, $row2, $cTime, $rowTime, $runDiff, $rowCPU, $rowTMM);
  
  # columns in 'raw_data' worksheet, NOT the 'summary' worksheet
  my %r = ('rowtime'      => 'A',
           'rowcpu'       => 'B',
           'rowtmm'       => 'C',
           'memutil'      => 'D',
           'cltBytesIn'   => 'E',
           'cltBytesOut'  => 'F',
           'cltPktsIn'    => 'G',
           'cltPktsOut'   => 'H',
           'svrBytesIn'   => 'I',
           'svrBytesOut'  => 'J',
           'svrPktsIn'    => 'K',
           'svrPktsOut'   => 'L',
           'cltTotConns'  => 'N',
           'svrTotConns'  => 'P',
           'httpRequests' => 'Q',
          );


  for ($row0 = 1; $row0 < $numRows; $row0++) {
    $row1    = $row0+1;
    $row2    = $row0+2;

    $cTime   = 'raw_data!'.$r{'rowtime'}.$row2.'-raw_data!'.$r{'rowtime'}.$row1;

    # splitting these out is required so a different format can be applied to numbers
    $rowTime = '=raw_data!'.$r{'rowtime'}.$row2;
    $rowCPU  = '=raw_data!'.$r{'rowcpu'}.$row2;
    $rowTMM  = '=raw_data!'.$r{'rowtmm'}.$row2;
    $runDiff = '='.$cTime;

    # @rowData contains formulas required to populate the summary data sheet.
    # In order, they are: memutil, client bits/sec in, client bits/sec out,
    #                     server bits/sec in, server bits/sec out, client conns/sec,
    #                     server conns/sec, http requests/sec
    @rowData = (
      '=raw_data!'   .$r{'memutil'}.$row2,
      '=(((raw_data!'.$r{'cltBytesIn'} .$row2.'-raw_data!'.$r{'cltBytesIn'} .$row1.')/('.$cTime.'))*8)',
      '=(((raw_data!'.$r{'cltBytesOut'}.$row2.'-raw_data!'.$r{'cltBytesOut'}.$row1.')/('.$cTime.'))*8)',
      '=(((raw_data!'.$r{'svrBytesIn'} .$row2.'-raw_data!'.$r{'svrBytesIn'} .$row1.')/('.$cTime.'))*8)',
      '=(((raw_data!'.$r{'svrBytesOut'}.$row2.'-raw_data!'.$r{'svrBytesOut'}.$row1.')/('.$cTime.'))*8)',
      '=((raw_data!'.$r{'cltPktsIn'} .$row2.'-raw_data!'.$r{'cltPktsIn'} .$row1.')/('.$cTime.'))',
      '=((raw_data!'.$r{'cltPktsOut'}.$row2.'-raw_data!'.$r{'cltPktsOut'}.$row1.')/('.$cTime.'))',
      '=((raw_data!'.$r{'svrPktsIn'} .$row2.'-raw_data!'.$r{'svrPktsIn'} .$row1.')/('.$cTime.'))',
      '=((raw_data!'.$r{'svrPktsOut'}.$row2.'-raw_data!'.$r{'svrPktsOut'}.$row1.')/('.$cTime.'))',
      '=((raw_data!' .$r{'cltTotConns'}.$row2.'-raw_data!'.$r{'cltTotConns'}.$row1.')/('.$cTime.'))',
      '=((raw_data!' .$r{'svrTotConns'}.$row2.'-raw_data!'.$r{'svrTotConns'}.$row1.')/('.$cTime.'))',
      '=((raw_data!' .$r{'httpRequests'}.$row2.'-raw_data!'.$r{'httpRequests'}.$row1.')/('.$cTime.'))',
    );

    $DEBUG && print Dumper(\@rowData);
    $worksheet->write($row0, 0, [$rowTime, $runDiff], ${$formats}{'decimal4'});
    $worksheet->write($row0, 2, $rowCPU,   ${$formats}{'decimal2'});
    $worksheet->write($row0, 3, $rowTMM,   ${$formats}{'decimal2'});
    $worksheet->write($row0, 4, \@rowData, ${$formats}{'standard'});
  }
}

sub write_chartData() {
  my $worksheet = shift;
  my $formats   = shift;
  my $numRows   = shift;
  my ($row0, $row1, $row2, $col, $cTime, $rowTime, $runDiff, $rowCPU, $rowTMM);

  # columns in 'raw_data' worksheet
  my %r = ('rowtime'      => 'A',
           'rowcpu'       => 'B',
           'rowtmm'       => 'C',
           'memutil'      => 'D',
           'cltBytesIn'   => 'E',
           'cltBytesOut'  => 'F',
           'svrBytesIn'   => 'I',
           'svrBytesOut'  => 'J',
           'cltTotConns'  => 'N',
           'cltCurConns'  => 'M',
           'svrCurConns'  => 'O',
           'svrTotConns'  => 'P',
           'httpRequests' => 'Q',
          );
  # columns in 'summary' worksheet
  #my %s = ('cltConnRate'  => 'N',
  #         'srvConnRate'  => 'O',
  #         'httpReqRate'  => 'P',
  #        );

  for ($row0 = 1; $row0 < $numRows; $row0++) {
    $row1    = $row0+1;
    $row2    = $row0+2;
    $cTime   = 'raw_data!'.$r{'rowtime'}.$row2.'-raw_data!'.$r{'rowtime'}.$row1;

    # splitting these out is required so different formats can be applied
    $rowTime = '=raw_data!'.$r{'rowtime'}.$row2;
    $rowCPU  = '=raw_data!'.$r{'rowcpu'}.$row2;
    $rowTMM  = '=raw_data!'.$r{'rowtmm'}.$row2;
    $runDiff = '='.$cTime;
   
    # @rowData contains formulas required to populate the chart_data worksheet
    # All calculations are performed on the values in the 'raw_data' worksheet
    # In order, they are: memutil, client bits/sec in, client bits/sec out,
    #                     server bits/sec in, server bits/sec out, client conns/sec,
    #                     server conns/sec
    @rowData = (
      '=(raw_data!'   .$r{'memutil'}    .$row2.'/'.MB.')',
      '=((((raw_data!'.$r{'cltBytesIn'} .$row2.'-raw_data!'.$r{'cltBytesIn'} .$row1.')/('.$cTime.'))*8)/1000000)',
      '=((((raw_data!'.$r{'cltBytesOut'}.$row2.'-raw_data!'.$r{'cltBytesOut'}.$row1.')/('.$cTime.'))*8)/1000000)',
      '=((((raw_data!'.$r{'svrBytesIn'} .$row2.'-raw_data!'.$r{'svrBytesIn'} .$row1.')/('.$cTime.'))*8)/1000000)',
      '=((((raw_data!'.$r{'svrBytesOut'}.$row2.'-raw_data!'.$r{'svrBytesOut'}.$row1.')/('.$cTime.'))*8)/1000000)',
      '=raw_data!'   .$r{'cltCurConns'}.$row2,
      '=raw_data!'   .$r{'svrCurConns'}.$row2,
      '=((raw_data!' .$r{'cltTotConns'}.$row2.'-raw_data!'.$r{'cltTotConns'}.$row1.')/('.$cTime.'))',
      '=((raw_data!' .$r{'svrTotConns'}.$row2.'-raw_data!'.$r{'svrTotConns'}.$row1.')/('.$cTime.'))',
      '=((raw_data!' .$r{'httpRequests'}.$row2.'-raw_data!'.$r{'httpRequests'}.$row1.')/('.$cTime.'))',
      '=raw_data!'   .$r{'cltCurConns'}.$row2.'+raw_data!'.$r{'svrCurConns'}.$row2,

    );
      # These lines were replaced with direct references to 'raw_data' rather than the 'Summary' worksheet.
      # TODO: Removed the Summary worksheet (Jesse, 20140919)
      #'=summary!'    .$s{'cltConnRate'}.$row2,  # K, clt conns/sec
      #'=summary!'    .$s{'srvConnRate'}.$row2,  # L, svr conns/sec
      #'=summary!'    .$s{'httpReqRate'}.$row2,  # M, requests/sec

    $DEBUG && print Dumper(\@rowData);
    $worksheet->write($row0, 0, $rowTime,  ${$formats}{'decimal0'});
    $worksheet->write($row0, 1, $rowCPU,   ${$formats}{'decimal0'});
    $worksheet->write($row0, 2, $rowTMM,   ${$formats}{'decimal0'});
    $worksheet->write($row0, 3, \@rowData, ${$formats}{'standard'});
  }
}

## returns a has containing the data-oids
sub get_f5_oids() {
  my %oidlist = (
      'ssCpuRawUser'            => '.1.3.6.1.4.1.2021.11.50.0',
      'ssCpuRawNice'            => '.1.3.6.1.4.1.2021.11.51.0',
      'ssCpuRawSystem'          => '.1.3.6.1.4.1.2021.11.52.0',
      'ssCpuRawIdle'            => '.1.3.6.1.4.1.2021.11.53.0',
      'tmmTotalCycles'          => '.1.3.6.1.4.1.3375.2.1.1.2.1.41.0',
      'tmmIdleCycles'           => '.1.3.6.1.4.1.3375.2.1.1.2.1.42.0',
      'tmmSleepCycles'          => '.1.3.6.1.4.1.3375.2.1.1.2.1.43.0',
      'tmmTotalMemoryUsed'      => '.1.3.6.1.4.1.3375.2.1.1.2.1.45.0',
      'tmmClientBytesIn'        => '.1.3.6.1.4.1.3375.2.1.1.2.1.3.0',
      'tmmClientBytesOut'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.5.0',
      'tmmClientPktsIn'         => '.1.3.6.1.4.1.3375.2.1.1.2.1.2.0',
      'tmmClientPktsOut'        => '.1.3.6.1.4.1.3375.2.1.1.2.1.4.0',
      'tmmClientTotConns'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.7.0',
      'tmmClientCurConns'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.8.0',
      'tmmServerBytesIn'        => '.1.3.6.1.4.1.3375.2.1.1.2.1.10.0',
      'tmmServerBytesOut'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.12.0',
      'tmmServerPktsIn'         => '.1.3.6.1.4.1.3375.2.1.1.2.1.9.0',
      'tmmServerPktsOut'        => '.1.3.6.1.4.1.3375.2.1.1.2.1.11.0',
      'tmmServerTotConns'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.14.0',
      'tmmServerCurConns'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.15.0',
      'sysStatHttpRequests'     => '.1.3.6.1.4.1.3375.2.1.1.2.1.56.0',
      'pvaClientPktsIn'         => '.1.3.6.1.4.1.3375.2.1.1.2.1.16.0',
      'pvaClientBytesIn'        => '.1.3.6.1.4.1.3375.2.1.1.2.1.17.0',
      'pvaClientPktsOut'        => '.1.3.6.1.4.1.3375.2.1.1.2.1.18.0',
      'pvaClientBytesOut'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.19.0',
      'pvaServerPktsIn'         => '.1.3.6.1.4.1.3375.2.1.1.2.1.23.0',
      'pvaServerBytesIn'        => '.1.3.6.1.4.1.3375.2.1.1.2.1.24.0',
      'pvaServerPktsOut'        => '.1.3.6.1.4.1.3375.2.1.1.2.1.25.0',
      'pvaServerBytesOut'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.26.0',
      'pvaClientTotConns'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.21.0',
      'pvaClientCurConns'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.22.0',
      'pvaServerTotConns'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.28.0',
      'pvaServerCurConns'       => '.1.3.6.1.4.1.3375.2.1.1.2.1.29.0',
  );
      #'pvaClientPktsIn'         => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.2.3.48.46.48',
      #'pvaClientBytesIn'        => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.3.3.48.46.48',
      #'pvaClientPktsOut'        => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.4.3.48.46.48',
      #'pvaClientBytesOut'       => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.5.3.48.46.48',
      #'pvaServerPktsIn'         => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.9.3.48.46.48',
      #'pvaServerBytesIn'        => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.10.3.48.46.48',
      #'pvaServerPktsOut'        => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.11.3.48.46.48',
      #'pvaServerBytesOut'       => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.12.3.48.46.48',
      #'pvaClientTotConns'       => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.7.3.48.46.48',
      #'pvaClientCurConns'       => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.8.3.48.46.48',
      #'pvaServerTotConns'       => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.14.3.48.46.48',
      #'pvaServerCurConns'       => '.1.3.6.1.4.1.3375.2.1.8.1.3.1.15.3.48.46.48',
  return(%oidlist);
}

sub get_profile_oids() {
  my %profileOids = ( 'userStatProfile1'  => '1.3.6.1.4.1.3375.2.2.6.19.2.3.1',
  );

  return(%profileOids);
}

# returns a hash containing oids that will be polled only once
sub get_static_oids() {
  my %oidlist = ( 'ltmVersion'   => '.1.3.6.1.4.1.3375.2.1.4.2.0',
                  'ltmBuild'     => '.1.3.6.1.4.1.3375.2.1.4.3.0',
                  'platform'     => '.1.3.6.1.4.1.3375.2.1.3.3.1.0',
                  'cpuCount'     => '.1.3.6.1.4.1.3375.2.1.1.2.1.38.0',
                  'totalMemory'  => '.1.3.6.1.4.1.3375.2.1.1.2.1.44.0',
                  'hostName'     => '.1.3.6.1.4.1.3375.2.1.6.2.0',
                  'bladeCount'   => '.1.3.6.1.4.1.3375.2.1.7.4.1.0',
                );

  return(%oidlist);
}

# returns a has containing oids that track errors
sub get_err_oids() {
  my %oidlist = (
      'incomingPktErrors'   => '.1.3.6.1.4.1.3375.2.1.1.2.1.47.0',
      'outgoingPktErrors'   => '.1.3.6.1.4.1.3375.2.1.1.2.1.48.0',
      'IPDroppedPkts'       => '.1.3.6.1.4.1.3375.2.1.1.2.7.4.0',
      'vipNonSynDeny'       => '.1.3.6.1.4.1.3375.2.1.1.2.21.20.0',
      'cmpConnRedirected'   => '.1.3.6.1.4.1.3375.2.1.1.2.21.23.0',
      'connMemErrors'       => '.1.3.6.1.4.1.3375.2.1.1.2.21.24.0',
      'sysIpStatDropped'    => '.1.3.6.1.4.1.3375.2.1.1.2.7.4.0',
  );

  return(%oidlist);
}

sub mk_perf_xls() {
  my $fname   = shift;
  my $rawHdrs = shift;
  my $sumHdrs = shift;
  my $chtHdrs = shift;
  my $dutHdrs = shift;
  my %hdrfmts;

  ## create Excel workbook
  my $workbook = Excel::Writer::XLSX->new($fname);

  # define formatting
  $DEBUG && print "Generating workbook formats (mk_perf_xls())\n";
  $hdrfmts{'text'}     = $workbook->add_format(align => 'center');
  $hdrfmts{'headers'}  = $workbook->add_format(align => 'center', bold => 1, bottom => 1);
  $hdrfmts{'standard'} = $workbook->add_format(align => 'center', num_format => '#,##0');
  $hdrfmts{'decimal0'} = $workbook->add_format(align => 'center', num_format => '#,##0');
  $hdrfmts{'decimal1'} = $workbook->add_format(align => 'center', num_format => '0.0');
  $hdrfmts{'decimal2'} = $workbook->add_format(align => 'center', num_format => '0.00');
  $hdrfmts{'decimal4'} = $workbook->add_format(align => 'center', num_format => '0.0000');

  ## create worksheets
  # the 'charts' worksheet will contain graphs using data from the 'summary' sheet.
  my $charts = $workbook->add_worksheet('charts');
  $charts->hide_gridlines(2);
  $charts->set_zoom(100);
  $charts->set_column('A:A', 30);
  $charts->set_column('B:D', 10);
  $charts->set_column('E:E', 15);
  $charts->set_column('F:G', 10);
  $charts->activate();

  # create the worksheet that will contain the data used for the charts
  # similiar to the 'summary' worksheet, but with fewer columns
  my $chtData = $workbook->add_worksheet('chart_data');
  $chtData->set_zoom(80);
  $chtData->set_column('A:C', 9);
  $chtData->set_column('D:G', 15);
  $chtData->set_column('H:O', 18);
  #$chtData->activate();

  # the 'summary' worksheet contains summarized data from the 'raw_data' worksheet
  my $summary = $workbook->add_worksheet('summary');
  $summary->set_zoom(80);
  $summary->set_column('A:C', 9);
  $summary->set_column('D:D', 15);
  $summary->set_column('E:E', 13);
  $summary->set_column('F:Q', 18);

  # contains the raw data retrieved with SNMP
  my $rawdata = $workbook->add_worksheet('raw_data');
  $rawdata->set_zoom(80);
  $rawdata->set_column('A:C', 9);
  $rawdata->set_column('D:Z', 17);
  #$rawdata->activate();

  $charts->write( 0, 0, $dutHdrs, $hdrfmts{'headers'});
  $chtData->write(0, 0, $chtHdrs, $hdrfmts{'headers'});
  $summary->write(0, 0, $sumHdrs, $hdrfmts{'headers'});
  $rawdata->write(0, 0, $rawHdrs, $hdrfmts{'headers'});

  return($workbook, $rawdata, $summary, $chtData, $charts, %hdrfmts);
}

# Create the charts that will be displayed on the 'charts' sheet
sub mk_charts() {
  my $fname     = shift;
  my $worksheet = shift;
  my $numRows   = shift;

  $DEBUG && print "Writing 'charts' worksheet.\n";

  ## CPU Usage chart
  my $chtCpu  = $fname->add_chart( type => 'line', embedded => 1);
  $chtCpu->set_title ( name => 'CPU Utilization', name_font => { size => 12, bold => 0} );
  $chtCpu->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => 45 } );
  $chtCpu->set_y_axis( name => 'CPU Usage', min => 0, max => 100 );
  $chtCpu->set_legend( position => 'bottom' );
  $chtCpu->add_series(
    name        => '=chart_data!$B$1',
    values      => '=chart_data!$B$2:$B$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'blue' },
    marker      => { type  => 'none' },
  );
  $chtCpu->add_series(
    name        => '=chart_data!$C$1',
    values      => '=chart_data!$C$2:$C$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'red' },
    marker      => { type  => 'none' },
  );
  $worksheet->insert_chart( 'A4', $chtCpu, 10, 0);

  ## Connection Rate
  my $chtCPS  = $fname->add_chart( type => 'line', embedded => 1);
  $chtCPS->set_title ( name => 'Connection Rate', name_font => { size => 12, bold => 0} );
  $chtCPS->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => 45 } );
  $chtCPS->set_y_axis( name => 'Connections/Second', min => 0);
  $chtCPS->set_legend( position => 'bottom' );
  $chtCPS->add_series(
    name        => '=chart_data!$K$1',
    values      => '=chart_data!$K$2:$K$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'blue' },
    marker      => { type  => 'none' },
  );
  $chtCPS->add_series(
    name        => '=chart_data!$L$1',
    values      => '=chart_data!$L$2:$L$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'red' },
    marker      => { type  => 'none' },
  );
  $worksheet->insert_chart( 'E4', $chtCPS, 50, 0);

  ## Throughput chart
  my $chtTput = $fname->add_chart( type => 'line', embedded => 1);
  $chtTput->set_title ( name => 'Client Throughput', name_font => { size => 12, bold => 0} );
  $chtTput->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => 45 } );
  $chtTput->set_y_axis( name => 'Throughput (Mbps)', min => 0);
  $chtTput->set_legend( position => 'bottom' );
  $chtTput->add_series(
    name        => '=chart_data!$E$1',
    values      => '=chart_data!$E$2:$E$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'blue' },
    marker      => { type  => 'none' },
  );
  $chtTput->add_series(
    name        => '=chart_data!$F$1',
    values      => '=chart_data!$F$2:$F$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'red' },
    marker      => { type  => 'none' },
  );
  $worksheet->insert_chart( 'A19', $chtTput, 10, 0);

  ## Transaction Rate
  my $chtTPS  = $fname->add_chart( type => 'line', embedded => 1);
  $chtTPS->set_title ( name => 'Transaction Rate', name_font => { size => 12, bold => 0} );
  $chtTPS->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => 45 } );
  $chtTPS->set_y_axis( name => 'Transactions/Second', min => 0);
  $chtTPS->set_legend( position => 'bottom' );
  $chtTPS->add_series(  # HTTP Transaction rate
    name        => '=chart_data!$M$1',
    values      => '=chart_data!$M$2:$M$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'green' },
    marker      => { type  => 'none' },
  );
  $chtTPS->add_series(  # Client-side connection rate
    name        => '=chart_data!$K$1',
    values      => '=chart_data!$K$2:$K$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'blue' },
    marker      => { type  => 'none' },
  );
  $chtTPS->add_series(  # Server-side connection rate
    name        => '=chart_data!$L$1',
    values      => '=chart_data!$L$2:$L$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'red' },
    marker      => { type  => 'none' },
  );
  $worksheet->insert_chart( 'E19', $chtTPS, 50, 0);

  ## Memory usage chart
  my $chtMem  = $fname->add_chart( type => 'line', embedded => 1);
  $chtMem->set_title ( name => 'Memory Utilization', name_font => { size => 12, bold => 0} );
  $chtMem->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => 45 } );
  $chtMem->set_y_axis( name => 'Memory Usage (MB)', min => 0);
  $chtMem->set_legend( position => 'none' );
  $chtMem->add_series(
    name        => '=chart_data!$D$1',
    values      => '=chart_data!$D$2:$D$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'blue' },
    marker      => { type  => 'none' },
  );
  $worksheet->insert_chart( 'A34', $chtMem, 10, 0);

  ## Concurrency
  my $chtCC   = $fname->add_chart( type => 'line', embedded => 1);
  $chtCC->set_title ( name => 'Concurrency', name_font => { size => 12, bold => 0} );
  $chtCC->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => 45 } );
  $chtCC->set_y_axis( name => 'Concurrent Connections', min => 0);
  $chtCC->set_legend( position => 'bottom' );
  $chtCC->add_series(
    name        => '=chart_data!$I$1',
    values      => '=chart_data!$I$2:$I$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'blue' },
    marker      => { type  => 'none' },
  );
  $chtCC->add_series(
    name        => '=chart_data!$J$1',
    values      => '=chart_data!$J$2:$J$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'red' },
    marker      => { type  => 'none' },
  );
  $chtCC->add_series(
    name        => '=chart_data!$N$1',
    values      => '=chart_data!$N$2:$N$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'green' },
    marker      => { type  => 'none' },
  );
  $worksheet->insert_chart( 'E34', $chtCC, 50, 0);

  return(1);
}

# Close the spreadsheet -- REQUIRED
sub close_xls() {
  my $xls = shift;
  $xls->close();
  return(1);
}

# Serialize and write out the json blob
sub json_fwrite() {
  open(JSONOUT, ">", $jsonName) or die "Could not open $jsonName for writing.\n";
  print JSONOUT encode_json(\%json_buffer);
  close(JSONOUT);
}

# saves the data when exitting the program immediately (CTRL+c)
sub exit_now() {
  if ($JSONOUT == 1 && $row > 0) {
    &json_fwrite();
  }
  if ($XLSXOUT == 1 && $row > 0) {
    print "\nStatistics collection cancelled. Attempting to save data.\n";
    #&write_summary($summary, \%formats, $row);
    &write_chartData($chtdata, \%formats, $row);
    &mk_charts($workbook, $charts, $row) if $row > 0;
    $workbook->close();
  }
  elsif ($XLSXOUT) {
    print "\nStatistics collection cancelled, no data collected.\n";
    $workbook->close();
  }
  exit(5);
}



# print script usage and exit with the supplied status
sub usage() {
  my $code = shift;

  print <<END;
  USAGE:  $0 -d <host> -l <total test length> -o <output file>
          $0 -h

  -d      IP or hostname to query (REQUIRED)
  -m      IP or hostname to monitor for the start of the test. Use this to 
          monitor the active LTM in a failover pair, but record data from the 
          standby LTM.
  -l      Full Test duration              (default: 130 seconds)
  -i      Seconds between polling cycles  (default: 10 seconds)
  -o      XLSX output filename            (default: /dev/null)
  -j      JSON output filename            (default: /dev/null)
  -p      Pause time; the amount of time to wait following the start of the 
          test before beginning polling. Should match the ramp-up time (default: 0)
  -v      Verbose output (print stats)
  -B      Bypass start-of-test detection and start polling immediately
  -h      Print usage and exit

END

  exit($code);
}
