#! /usr/bin/env perl

#####################################################
# Collects numerous statistics from a target BIP-IP while under test
# writes the output to an excel spreadsheet
#
# Copyright, F5 Networks, 2009-2013
# Written by: Jesse Driskill, Product Management Engineer
#####################################################

## Required libraries
#
use warnings;
use strict;
use Config;
use Getopt::Std;
use Net::SNMP qw(:snmp);
import Time::HiRes qw(usleep time clock);
use Excel::Writer::XLSX;
use Data::Dumper;
#use Spreadsheet::WriteExcel;

$!++;  # No buffer for STDOUT
$SIG{'INT'} = \&exit_now;  # handle ^C nicely

## retrieve and process CLI paramters
#
our (%opts, $DATAOUT, $BYPASS, $DEBUG, $VERBOSE, $CUSTOMER, $TESTNAME, $COMMENTS);
our ($PAUSE);
getopts('d:m:l:o:C:T:m:c:s:i:p:BDvh', \%opts);

# print usage and exit
&usage(0) if $opts{'h'};

if (!$opts{'d'}) {
  warn("Must provide a hostname or IP address to query\n");
  &usage(1);
}

my $host      = $opts{'d'};                 # snmp host to poll
my $secondary = $opts{'m'};                 # monitoring host
my $testLen   = $opts{'l'} || 130;          # total duration of test in seconds
my $xlsName   = $opts{'o'} || '/dev/null';  # output file name
my $snmpVer   = $opts{'s'} || 'v2c';        # snmp version
my $comm      = $opts{'c'} || 'public';     # community string    
my $cycleTime = $opts{'i'} || 10;           # polling interval
my $pause     = $opts{'p'} || 0;            # pause time


## This commented block is probably unnecessary and should be removed (jesse, 2012-06-15)
#if ($opts{'D'}) { $DEBUG      = 1; }
#if ($opts{'o'}) { $DATAOUT    = 1; }
#if ($opts{'B'}) { $BYPASS     = 1; }
#if ($opts{'v'}) { $VERBOSE    = 1; }
#if ($opts{'C'}) { $CUSTOMER   = $opts{'C'}; }
#if ($opts{'T'}) { $TESTNAME   = $opts{'T'}; }
#if ($opts{'m'}) { $COMMENTS   = $opts{'m'}; }

# $DATAOUT must be defined or the signal handler will throw an error if 
# the output file isn't setup. Only a cosmetic issue, but irritating.
$DEBUG      = ($opts{'D'} ? 1 : 0);
$DATAOUT    = ($opts{'o'} ? 1 : 0);
$VERBOSE    = ($opts{'v'} ? 1 : 0);
$BYPASS     = ($opts{'B'} ? 1 : 0);
$CUSTOMER   = ($opts{'C'} ? 1 : 0);
$TESTNAME   = ($opts{'T'} ? 1 : 0);
$COMMENTS   = ($opts{'m'} ? 1 : 0);

## normal vars
#
my $elapsed   = 0;       # total time test has been running
my %pollTimer = ();      # tracks the amount on time required for each poll operation


if ($DEBUG) {
  print Dumper(\%opts);
  #($_ = <<EOF) =~ s/^\s+//gm;
  print "
  DATAOUT:  $DATAOUT
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
my $usCycleTime  = $cycleTime * 1000000;

##
## Initialization and environment check
##

$VERBOSE && print("Host: ".$host."\nDuration: ".$testLen." seconds\nPolling Interval: ".
                    $cycleTime."\nFile: ".$xlsName."\n\n");

# Build the oid lists and varbind arrays
my (@dataList, @errorList, @staticList, @rowData);
my ($clientCurConns, $clientTotConns, $serverCurConns, $serverTotConns);
my ($cpuUsed, $cpuTicks, $cpuUtil, $cpuPercent, $tmmUtil, $tmmPercent);
my ($httpRequsts);
my ($memUsed, $hMem, $dataVals, $errorVals, $col);
my ($workbook, $summary, $raw_data, $chtdata, $charts);
my ($cBytesIn, $cBytesOut, $sBytesIn, $sBytesOut, $tBytesIn, $tBytesOut);
my ($cPktsIn, $cPktsOut, $sPktsIn, $sPktsOut);
my ($cNewConns, $sNewConns, $ccPktsIn, $ccPktsOut, $cBitsIn, $cBitsOut)   = (0, 0, 0, 0, 0, 0);
my ($row, $sBitsIn, $sBitsOut, $tBitsIn, $tBitsOut, $httpReq)             = (0, 0, 0, 0, 0, 0);
my ($slept, $sleepTime, $pollTime, $runTime, $lastLoopEnd, $loopTotal)    = (0, 0, 0, 0, 0, 0);
my $loopTime   = 1;

my (%formats, %xlsData);
my %staticOids  = &get_static_oids();
my %dataOids    = &get_f5_oids();
my %errorOids   = &get_err_oids();
my %oldData     = (ssCpuRawUser   => 0,
                   ssCpuRawNice   => 0,
                   ssCpuRawSystem => 0,
                   ssCpuRawIdle   => 0,
                   tmmTotalCycles => 0,
                   tmmIdleCycles  => 0,
                   tmmSleepCycles => 0,
                   cBytesIn       => 0,
                   cBytesOut      => 0,
                   sBytesIn       => 0,
                   sBytesOut      => 0,
                   httpReq        => 0,
                  );
my @dutInfoHdrs = qw(Host Platform Version Build Memory CPUs Blades);
my @chtDataHdrs = ('RunTime', 'SysCPU', 'TmmCPU', 'Memory', 'Clnt bitsIn/s', 
                   'Clnt bitsOut/s', 'Svr bitsIn/s', 'Svr bitsOut/s','Client CurConns',
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
                  );

while (my ($key, $value) = each(%staticOids)) { push(@staticList, $value); }
while (my ($key, $value) = each(%dataOids))   { push(@dataList, $value); }
while (my ($key, $value) = each(%errorOids))  { push(@errorList, $value); }

#my ($session, $error);
#if ( $snmpVer =~ m/v2c/ ) {
#  print "SNMP version v2c specified\n";
  my ($session, $error) = Net::SNMP->session(
    -hostname     => $host,
    -community    => $comm,
    -version      => $snmpVer,
    -maxmsgsize   => 8192,
    -nonblocking  => 0,
  );
  die($error."\n") if ($error);
#}
#elsif ($snmpVer =~ m/3/) {
#  print "SNMP version 3 specified\n";
#  my ($session, $error) = Net::SNMP->session(
#    -hostname     => $host,
#    -version      => 'snmpv3',
#    -username     => 'admin',
#    -authpassword => 'admin123',
#    -authProtocol => 'md5',
#    -maxmsgsize   => 8192,
#    -nonblocking  => 0,
#  );
#  die($error."\n") if ($error);
#}

# determine if logging is required and create the output files
if ($DATAOUT) {
  $DEBUG && print "Creating workbook ($xlsName)\n";
  ($workbook, $raw_data, $summary, $chtdata, $charts, %formats) = 
      &mk_perf_xls($xlsName, \@rawdataHdrs, \@summaryHdrs, \@chtDataHdrs, \@dutInfoHdrs);
}

# print out some information about the DUT being polled
my $result = $session->get_request( -varbindlist  => \@staticList);
print "Platform:    $result->{$staticOids{platform}}\n";
print "Memory:      $result->{$staticOids{totalMemory}} (".$result->{$staticOids{totalMemory}} / (1024*1024)." MB)\n";
print "# of CPUs:   $result->{$staticOids{cpuCount}}\n";
print "# of blades: $result->{$staticOids{bladeCount}}\n";
print "LTM Version: $result->{$staticOids{ltmVersion}}\n";
print "LTM Build:   $result->{$staticOids{ltmBuild}}\n";

# If a real xls is being written to, record DUT vital info on the first sheet
if ($xlsName !~ '/dev/null') {
  while (my ($k, $v) = each(%staticOids)) {
    print $k.": ".$result->{$v}."\n";
  }
  $charts->write("A2", $result->{$staticOids{hostName}},    $formats{text});
  $charts->write("B2", $result->{$staticOids{platform}},    $formats{text});
  $charts->write("C2", $result->{$staticOids{ltmVersion}},  $formats{text});
  $charts->write("D2", $result->{$staticOids{ltmBuild}},    $formats{text});
  $charts->write("E2", $result->{$staticOids{totalMemory}}, $formats{decimal0});
  $charts->write("F2", $result->{$staticOids{cpuCount}},    $formats{text});
  $charts->write("G2", $result->{$staticOids{bladeCount}},  $formats{text});
}


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
$pollTimer{'testStart'} = Time::HiRes::time();

while ($elapsed <= $testLen) {
  $pollTimer{activeStart} = Time::HiRes::time();


  # get snmp stats from DUT
  $dataVals = $session->get_request( -varbindlist  => \@dataList);
  die($session->error."\n") if (!defined($dataVals));

  $pollTimer{pollTime} = Time::HiRes::time();

  # update $runTime now so it will be accurate when written to the file
  $runTime    = sprintf("%.7f", ($pollTimer{pollTime} - $pollTimer{testStart}));

  # Get the exact time since the previous loop ended 
  # This is used to get an accurate value for the 'rate' counters
  ##Deprecated: $loopTime   = $pollTimer{pollTime} - $lastLoopEnd if ($lastLoopEnd);
  if ($lastLoopEnd) {
    $loopTime = $pollTimer{pollTime} - $lastLoopEnd;
    $elapsed += $loopTime;
  } else {
    $loopTime = 0;
  }

  # Before any real processing, remove any non-numeric values (i.e. 'noSuchInstance')
  for my $n (keys(%dataOids)) {
    if ($dataVals->{$dataOids{$n}} =~ /\D+/) {
      $xlsData{$n} = 1;
    }
    else {
      $xlsData{$n} = $dataVals->{$dataOids{$n}};
    }
  }

  # calculate accurate values from counters and record results
  ($cpuUtil, $cpuPercent) = &get_cpu_usage($dataVals, \%oldData);
  ($tmmUtil, $tmmPercent) = &get_tmm_usage($dataVals, \%oldData);

  $memUsed    = $dataVals->{$dataOids{tmmTotalMemoryUsed}};
  $hMem       = sprintf("%d", $memUsed / MB); # get Memory usage in MB

  # client and server current connections
  $clientCurConns = $dataVals->{$dataOids{tmmClientCurConns}}; # + $dataVals->{$dataOids{pvaClientCurConns}};
  $serverCurConns = $dataVals->{$dataOids{tmmServerCurConns}}; # + $dataVals->{$dataOids{pvaServerCurConns}};
  # total connections
  $clientTotConns = $dataVals->{$dataOids{tmmClientTotConns}}; # + $dataVals->{$dataOids{pvaClientTotConns}};
  $serverTotConns = $dataVals->{$dataOids{tmmServerTotConns}}; # + $dataVals->{$dataOids{pvaServerTotConns}};
  # Throughput
  $cBytesIn       = $dataVals->{$dataOids{tmmClientBytesIn}}  + $dataVals->{$dataOids{pvaClientBytesIn}};
  $cBytesOut      = $dataVals->{$dataOids{tmmClientBytesOut}} + $dataVals->{$dataOids{pvaClientBytesOut}};
  $sBytesIn       = $dataVals->{$dataOids{tmmServerBytesIn}}  + $dataVals->{$dataOids{pvaServerBytesIn}};
  $sBytesOut      = $dataVals->{$dataOids{tmmServerBytesOut}} + $dataVals->{$dataOids{pvaServerBytesOut}};
  # Packet rates
  $cPktsIn        = $dataVals->{$dataOids{tmmClientPktsIn}}   + $dataVals->{$dataOids{pvaClientPktsIn}};
  $cPktsOut       = $dataVals->{$dataOids{tmmClientPktsOut}}  + $dataVals->{$dataOids{pvaClientPktsOut}};
  $sPktsIn        = $dataVals->{$dataOids{tmmServerPktsIn}}   + $dataVals->{$dataOids{pvaServerPktsIn}};
  $sPktsOut       = $dataVals->{$dataOids{tmmServerPktsOut}}  + $dataVals->{$dataOids{pvaServerPktsOut}};

  if ($elapsed) { 
    # pre-format some vars
    $httpReq    = sprintf("%.0f", ($dataVals->{$dataOids{sysStatHttpRequests}} - $oldData{httpReq}) / $loopTime);
    $cNewConns  = sprintf("%.0f", ($dataVals->{$dataOids{tmmClientTotConns}} - $oldData{cTotConns}) / $loopTime);
    $sNewConns  = sprintf("%.0f", ($dataVals->{$dataOids{tmmServerTotConns}} - $oldData{sTotConns}) / $loopTime);
    $cBitsIn    = sprintf("%.0f", (($cBytesIn  - $oldData{cBytesIn})  * 8) / $loopTime / 1000000);
    $cBitsOut   = sprintf("%.0f", (($cBytesOut - $oldData{cBytesOut}) * 8) / $loopTime / 1000000);
    $sBitsIn    = sprintf("%.0f", (($sBytesIn  - $oldData{sBytesIn})  * 8) / $loopTime / 1000000);
    $sBitsOut   = sprintf("%.0f", (($sBytesOut - $oldData{sBytesOut}) * 8) / $loopTime / 1000000);
    $tBitsIn    = $cBitsIn + $sBitsIn;
    $tBitsOut   = $cBitsOut + $sBitsOut;

    if ( $elapsed > 0) {
    # This 'format' displays the standard data

      format STDOUT_TOP =
 @>>>>>     @>>>   @>>>    @>>>>>>>>     @>>>>>     @>>>>>  @>>>>>>>>  @>>>>>>>>>  @>>>>>>>>>  @>>>>>>  @>>>>>>>
"Time", "CPU", "TMM", "Mem (MB)", "C-CPS", "S-CPS", "HTTP_req", "Client CC", "Server CC", "In/Mbs", "Out/Mbs"
.

      format =
@####.###  @##.## @##.##    @#######  @########  @########  @####### @#########  @#########    @#####    @#####
$elapsed, $cpuUtil, $tmmUtil, $hMem, $cNewConns, $sNewConns, $httpReq, $clientCurConns, $serverCurConns, $cBitsIn, $cBitsOut
.
      write;
    }


    # If requested, write the output file.
    if ($DATAOUT) {
      $row++;
      $raw_data->write($row, 0, $runTime, $formats{decimal4});
      $raw_data->write($row, 1, $cpuUtil, $formats{decimal2});
      $raw_data->write($row, 2, $tmmUtil, $formats{decimal2});

      $raw_data->write( $row, 
                        3,
                        [$memUsed,
                        sprintf("%.0f", $cBytesIn),
                        sprintf("%.0f", $cBytesOut),
                        sprintf("%.0f", $cPktsIn),
                        sprintf("%.0f", $cPktsOut),
                        sprintf("%.0f", $sBytesIn),
                        sprintf("%.0f", $sBytesOut),
                        sprintf("%.0f", $sPktsIn),
                        sprintf("%.0f", $sPktsOut),
                        sprintf("%.0f", $clientCurConns),
                        sprintf("%.0f", $clientTotConns),
                        sprintf("%.0f", $serverCurConns),
                        sprintf("%.0f", $serverTotConns),
                        sprintf("%.0f", $dataVals->{$dataOids{sysStatHttpRequests}})],
                       $formats{'standard'});
    }
  }  

  # update 'old' data with the current values to calculate delta next cycle
  $oldData{ssCpuRawUser}   = $dataVals->{$dataOids{ssCpuRawUser}};
  $oldData{ssCpuRawNice}   = $dataVals->{$dataOids{ssCpuRawNice}};
  $oldData{ssCpuRawSystem} = $dataVals->{$dataOids{ssCpuRawSystem}};
  $oldData{ssCpuRawIdle}   = $dataVals->{$dataOids{ssCpuRawIdle}};
  $oldData{tmmTotalCycles} = $dataVals->{$dataOids{tmmTotalCycles}};
  $oldData{tmmIdleCycles}  = $dataVals->{$dataOids{tmmIdleCycles}};
  $oldData{tmmSleepCycles} = $dataVals->{$dataOids{tmmSleepCycles}};
  $oldData{cBytesIn}       = $cBytesIn;
  $oldData{cBytesOut}      = $cBytesOut;
  $oldData{sBytesIn}       = $sBytesIn;
  $oldData{sBytesOut}      = $sBytesOut;
  $oldData{sPktsIn}        = $sPktsIn;
  $oldData{sPktsOut}       = $sPktsOut;
  $oldData{cPktsIn}        = $cPktsIn;
  $oldData{cPktsOut}       = $cPktsOut;
  $oldData{cTotConns}      = $clientTotConns;
  $oldData{sTotConns}      = $serverTotConns;
  $oldData{httpReq}        = $dataVals->{$dataOids{sysStatHttpRequests}};
  #$oldData{cBytesIn}       = $dataVals->{$dataOids{tmmClientBytesIn}}  + $dataVals->{$dataOids{pvaClientBytesIn}};
  #$oldData{cBytesOut}      = $dataVals->{$dataOids{tmmClientBytesOut}} + $dataVals->{$dataOids{pvaClientBytesOut}};
  #$oldData{sBytesIn}       = $dataVals->{$dataOids{tmmServerBytesIn}}  + $dataVals->{$dataOids{pvaServerBytesIn}};
  #$oldData{sBytesOut}      = $dataVals->{$dataOids{tmmServerBytesOut}} + $dataVals->{$dataOids{pvaServerBytesOut}};
  #$oldData{sPktsIn}        = $dataVals->{$dataOids{tmmServerPktsIn}}   + $dataVals->{$dataOids{pvaServerPktsIn}};
  #$oldData{sPktsOut}       = $dataVals->{$dataOids{tmmServerPktsOut}}  + $dataVals->{$dataOids{pvaServerPktsOut}};
  #$oldData{cPktsIn}        = $dataVals->{$dataOids{tmmClientPktsIn}}   + $dataVals->{$dataOids{pvaClientPktsIn}};
  #$oldData{cPktsOut}       = $dataVals->{$dataOids{tmmClientPktsOut}}  + $dataVals->{$dataOids{pvaClientPktsOut}};

  if ($DEBUG) {
    format STDERR_TOP =
@>>>>>>>>  @||||||||||||  @||||||||||||||||| @||||||||||||||||| @||||||||||||||||| @|||||||||  @||||||||||||
"RunTime", "Elapsed", "activeStart", "activeStop", "pollTime", "loopTime", "lastLoopTotal"
.

    format STDERR =
@########  @######.#####  @###########.##### @###########.##### @###########.#####  @##.#####   @##.#####
$runTime, $elapsed, $pollTimer{activeStart}, $pollTimer{activeStop}, $pollTimer{pollTime}, $loopTime, $loopTotal
.
    write STDERR;
  }

  # Calculate how much time this polling cycle has required to determine how
  # long we should sleep before beginning the next cycle
  $pollTimer{activeStop} = Time::HiRes::time();
  $loopTotal = $pollTimer{'activeStop'} - $lastLoopEnd;
  $sleepTime = $cycleTime;

  $lastLoopEnd = Time::HiRes::time();
  Time::HiRes::sleep($sleepTime);
} 


if ($DATAOUT) {
  $DEBUG && print "Writing summary, chart_data, and charts worksheets...\n";
  # polling is now complete, time to write the summary formulas 
  &write_summary($summary, \%formats, $row);
  &write_chartData($chtdata, \%formats, $row);
  &mk_charts($workbook, $charts, $row);

  # close the workbook; required for the workbook to be usable.
  &close_xls($workbook);
}



##
## Subs
##

# delay the start of the script until the test is detected through pkts/sec
sub detect_test() {
  my $snmp = shift;
  my $oids = shift;
  my $pkts = 500;

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

    #$cTime   = 'raw_data!'.$r{'rowtime'}.$row2.'-raw_data!'.$r{'rowtime'}.$row1;
    $cTime   = $r{'rowtime'}.$row2.'-'.$r{'rowtime'}.$row1;

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
  my ($row0, $col, $row1, $row2, $cTime, $rowTime, $runDiff, $rowCPU, $rowTMM);

  # columns in 'raw_data' worksheet
  my %r = ('rowtime'      => 'A',
           'rowcpu'       => 'B',
           'rowtmm'       => 'C',
           'memutil'      => 'D',
           'cltBytesIn'   => 'E',
           'cltBytesOut'  => 'F',
           'svrBytesIn'   => 'I',
           'svrBytesOut'  => 'J',
           'cltCurConns'  => 'M',
           'svrCurConns'  => 'O',
           'httpRequests' => 'Q',
          );
  # columns in 'summary' worksheet
  my %s = ('cltConnRate'  => 'N',
           'srvConnRate'  => 'O',
           'httpReqRate'  => 'P',
          );

  for ($row0 = 1; $row0 < $numRows; $row0++) {
    $row1    = $row0+1;
    $row2    = $row0+2;
    $cTime   = $r{'rowtime'}.$row2.'-'.$r{'rowtime'}.$row1;

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
      '=raw_data!'   .$r{'memutil'}    .$row2,
      '=(((raw_data!'.$r{'cltBytesIn'} .$row2.'-raw_data!'.$r{'cltBytesIn'} .$row1.')/('.$cTime.'))*8)',
      '=(((raw_data!'.$r{'cltBytesOut'}.$row2.'-raw_data!'.$r{'cltBytesOut'}.$row1.')/('.$cTime.'))*8)',
      '=(((raw_data!'.$r{'svrBytesIn'} .$row2.'-raw_data!'.$r{'svrBytesIn'} .$row1.')/('.$cTime.'))*8)',
      '=(((raw_data!'.$r{'svrBytesOut'}.$row2.'-raw_data!'.$r{'svrBytesOut'}.$row1.')/('.$cTime.'))*8)',
      '=raw_data!'   .$r{'cltCurConns'}.$row2,
      '=raw_data!'   .$r{'svrCurConns'}.$row2,
      '=summary!'    .$s{'cltConnRate'}.$row2,
      '=summary!'    .$s{'srvConnRate'}.$row2,
      '=summary!'    .$s{'httpReqRate'}.$row2,
      '=raw_data!'   .$r{'cltCurConns'}.$row2.'+raw_data!'.$r{'svrCurConns'}.$row2,

    );

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

# Return System CPU utilization
sub get_cpu_usage() {
  my $curData = shift;
  my $oldData = shift;
  my ($cpuTotal, $cpuRaw, $cpuNice, $cpuSystem, $cpuIdle, $cpuUtil, $cpuPercent);
  my ($cpuUserDelta, $cpuNiceDelta, $cpuIdleDelta, $cpuSysDelta);

  $cpuUserDelta = $curData->{$dataOids{'ssCpuRawUser'}}   - $oldData->{'ssCpuRawUser'};
  $cpuNiceDelta = $curData->{$dataOids{'ssCpuRawNice'}}   - $oldData->{'ssCpuRawNice'};
  $cpuIdleDelta = $curData->{$dataOids{'ssCpuRawIdle'}}   - $oldData->{'ssCpuRawIdle'};
  $cpuSysDelta  = $curData->{$dataOids{'ssCpuRawSystem'}} - $oldData->{'ssCpuRawSystem'};
  $cpuTotal     = $cpuUserDelta + $cpuNiceDelta + $cpuIdleDelta + $cpuSysDelta;

  $cpuUtil      = (($cpuTotal - $cpuIdleDelta) / $cpuTotal) * 100;
  $cpuPercent   = sprintf("%.2f", ((($cpuTotal - $cpuIdleDelta) / $cpuTotal) * 100));

  return($cpuUtil, $cpuPercent);
}

# Return TMM CPU utilization
sub get_tmm_usage() {
  my $curData = shift;
  my $oldData = shift;
  my ($tmmTotal, $tmmIdle, $tmmSleep, $tmmUtil, $tmmPercent);

  $tmmTotal  = $curData->{$dataOids{'tmmTotalCycles'}} - $oldData->{'tmmTotalCycles'};
  $tmmIdle   = $curData->{$dataOids{'tmmIdleCycles'}}  - $oldData->{'tmmIdleCycles'};
  $tmmSleep  = $curData->{$dataOids{'tmmSleepCycles'}} - $oldData->{'tmmSleepCycles'};
  
  $tmmUtil    = (($tmmTotal - ($tmmIdle + $tmmSleep)) / $tmmTotal) * 100;
  $tmmPercent = sprintf("%.2f", $tmmUtil);

  return($tmmUtil, $tmmPercent);
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
  $chtCpu->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => -45 } );
  $chtCpu->set_y_axis( name => 'CPU Usage', min => 0, max => 100 );
  $chtCpu->set_legend( position => 'none' );
  $chtCpu->add_series(
    name        => '=chart_data!$B$1',
    values      => '=chart_data!$B$2:$B$'.($numRows-1),
    categories  => '=chart_data!$A$2:$A$'.($numRows-1),
    line        => { color => 'blue' },
    marker      => { type  => 'none' },
  );
  $worksheet->insert_chart( 'A4', $chtCpu, 10, 0);

  ## Connection Rate
  my $chtCPS  = $fname->add_chart( type => 'line', embedded => 1);
  $chtCPS->set_title ( name => 'Connection Rate', name_font => { size => 12, bold => 0} );
  $chtCPS->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => -45 } );
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
  $chtTput->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => -45 } );
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
  $chtTPS->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => -45 } );
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
  $chtMem->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => -45 } );
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
  $chtCC->set_x_axis( name => 'Time (Seconds)', num_font => { rotation => -45 } );
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

# saves the data when exitting the program immediately (CTRL+c)
sub exit_now() {
  if ($DATAOUT == 1 && $row > 0) {
    print "\nStatistics collection cancelled. Attempting to save data.\n";
    &write_summary($summary, \%formats, $row);
    &write_chartData($chtdata, \%formats, $row);
    &mk_charts($workbook, $charts, $row) if $row > 0;
    $workbook->close();
  }
  else {
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
  -l      Full Test duration             (default: 130 seconds)
  -i      Seconds between polling cycles (default: 10 seconds)
  -o      Output filename                (default: /dev/null)
  -p      Pause time; the amount of time to wait following the start of the 
          test before beginning polling. Should match the ramp-up time (default: 0)
  -v      Verbose output (print stats)
  -B      Bypass start-of-test detection and start polling immediately
  -h      Print usage and exit

END

  exit($code);
}

## Additional real-time output formats
#    format STDOUT_TOP =
# @>>>>>     @>>>   @>>>    @>>>>>>>>     @>>>>>     @>>>>>  @>>>>>>>>>  @>>>>>>>>>  @>>>>>>  @>>>>>>>
#"Time", "CPU", "TMM", "Mem (MB)", "C-CPS", "S-CPS", "Client CC", "Server CC", "In/Mbs", "Out/Mbs"
#.
#
#    format =
#@####.###  @##.## @##.##    @#######  @########  @########  @#########  @#########   @#####    @#####
#$elapsed, $cpuUtil, $tmmUtil, $hMem, $cNewConns, $sNewConns, $clientCurConns, $serverCurConns, $cBitsIn, $cBitsOut
#.
#     write;
#  }

## This 'format' displays the standard data, but substitutes packets/second for throughput
#    format STDOUT_TOP =
#@>>>>   @>>>  @>>>>  @>>>>>>>>>> @>>>>>> @>>>>>> @>>>>>>>>>>>> @>>>>>>>>>>>>  @>>>>>>>>   @>>>>>>>>
#"Time", "CPU", "TMM", "Memory (MB)", "C-CPS", "S-CPS", "In/Mbps", "Out/Mbps", "cPPS/in", "sPPS/in"
#.
#
#    format =
#@#### @##.## @##.## @>>>>>>>>>> @>>>>>> @>>>>>>     @>>>>>>>>     @>>>>>>>>   @>>>>>>>>   @>>>>>>>>
#$elapsed, $cpuUtil, $tmmUtil, $hMem, $cNewConns, $sNewConns, $cBitsIn, $cBitsOut, $cPktsIn, $sPktsIn 
#.
#    write;
#  }

## This 'format' emphasizes connections and PPS
#    format STDOUT_TOP =
#@>>>>   @>>>  @>>>>  @>>>>>>>>>> @>>>>>> @>>>>>> @>>>>>>>>>>>> @>>>>>>>>>>>>  @>>>>>>>>   @>>>>>>>>
#"Time", "CPU", "TMM", "Memory (MB)", "C-CPS", "S-CPS", "Client Conns", "Server Conns", "cPPS/in", "sPPS/in"
#.
#
#    format =
#@#### @##.## @##.## @>>>>>>>>>> @>>>>>> @>>>>>>     @>>>>>>>>     @>>>>>>>>   @>>>>>>>>   @>>>>>>>>
#$elapsed, $cpuUtil, $tmmUtil, $hMem, $cNewConns, $sNewConns, $clientCurConns, $serverCurConns, $cPktsIn, $sPktsIn
#.
#    write;
#  }

