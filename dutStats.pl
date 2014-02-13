#! /usr/bin/perl

#####################################################
# Collects numerous statistics from a target BIP-IP while under test, and
# writes the output to an excel spreadsheet
#
# Copyright, F5 Networks, 2008-2011
# Written by: Jesse Driskill, Product Management Engineer
#####################################################

## Required libraries
#
use warnings;
use strict;
use Config;
use Getopt::Std;
use Data::Dumper;
use Time::HiRes;

$!++;  # No buffer for STDOUT

## retrieve and process CLI paramters
#
our (%opts, $DATAOUT, $BYPASS, $DEBUG, $VERBOSE, $CUSTOMER, $TESTNAME, $COMMENTS);
our ($PAUSE);
getopts('d:l:o:C:T:m:c:s:i:p:BDvh', \%opts);


# print usage and exit
&usage(0) if $opts{'h'};

my $host      = $opts{'d'} || 'localhost';  # snmp host to poll
my $testLen   = $opts{'l'} || 86400;        # total duration of test in seconds
my $xlsName   = $opts{'o'} || '/dev/null';  # output file name
my $snmpVer   = $opts{'s'} || 'v2c';        # snmp version
my $comm      = $opts{'c'} || 'public';     # community string    
my $cycleTime = $opts{'i'} || 10;           # polling interval
my $pause     = $opts{'p'} || 0;            # pause time


if ($opts{'o'}) { $DATAOUT    = 1; }
if ($opts{'B'}) { $BYPASS     = 0; } else { $BYPASS = 1; }
if ($opts{'D'}) { $DEBUG      = 1; }
if ($opts{'v'}) { $VERBOSE    = 1; }
# Always enable verbose mode for dutStats
$VERBOSE  = 1;


## normal vars
#
my $elapsed   = 0;       # total time test has been running

# tracks the amount on time required for each poll operation
my %pollTimer = (
  'testStart'     => 0,
  'activeStart'   => 0,
  'activeStop'    => 0,
  'pollTime'      => 0,
  'activeRunTime' => 0,
);


if ($DEBUG) {
  print Dumper(\%opts);
  ($_ = <<EOF) =~ s/^\s+//gm;
  DATAOUT:  $DATAOUT
  BYPASS:   $BYPASS
  DEBUG:    $DEBUG
  VERBOSE:  $VERBOSE
  CUSTOMER: $CUSTOMER
  TESTNAME: $TESTNAME
  COMMENTS: $COMMENTS
EOF
}

# additional constants
use constant MB   => 1024*1024;
my $usCycleTime   = $cycleTime * 1000000;
my $sleepTime     = $cycleTime;

my %snmpOpts = ( 'host' => $host,
                 'comm' => $comm,
               );

##
## Initialization and environment check
##

$VERBOSE && print("Host: ".$host."\nDuration: ".$testLen." seconds\nPolling Interval: ".
                    $cycleTime."\nFile: ".$xlsName."\n\n");

# Build the oid lists and varbind arrays
my (@dataList, @errorList, @staticList, @rowData);
my ($clientCurConns, $clientTotConns, $serverCurConns, $serverTotConns);
my ($cpuUsed, $cpuTicks, $cpuUtil, $cpuPercent, $tmmUtil, $tmmPercent);
my ($memUsed, $hMem, $dataVals, $errorVals, $col, $row, $numRows);
my ($workbook, $summary, $raw_data, $chtdata, $charts);
my ($cBytesIn, $cBytesOut, $sBytesIn, $sBytesOut, $cPktsIn, $cPktsOut)  = (0, 0, 0, 0, 0, 0);
my ($cNewConns, $sNewConns, $ccPktsIn, $ccPktsOut, $cBitsIn, $cBitsOut) = (0, 0, 0, 0, 0, 0);
my ($slept, $loopTime, $pollTime, $runTime, $lastLoopEnd, $loopTotal)   = (0, 0, 0, 0, 0, 0);

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
                  );
my @dutInfoHdrs = qw(Host Platform Version Build Memory CPUs Blades);
my @rawdataHdrs = ('RunTime', 'SysCPU', 'TmmCPU', 'Memory', 'Client bytesIn', 'Client bytesOut', 
                   'Client pktsIn', 'Client pktsOut', 'Server btyesIn', 'Server bytesOut', 
                   'Server pktsIn', 'Server pktsOut', 'Client curConns', 'Client totConns', 
                   'Server curConns', 'Server totConns',
                  );

while (my ($key, $value) = each(%staticOids)) { push(@staticList, $value); }
while (my ($key, $value) = each(%dataOids))   { push(@dataList, $value); }


# TODO: Write a sub that starts a csv by writing the first line (column headers)
# determine if logging is required and create the output files
#if ($DATAOUT) {
#  $DEBUG && print "Creating csv...";
#  $csv = create_csv(\@rawdataHdrs)
#}

# TODO: Write the "get_platform_info()" sub
# print out some information about the DUT being polled
my $result = &get_snmp_data(\%staticOids, \%snmpOpts);
print "Platform:    $result->{$staticOids{platform}}\n";
print "Memory:      ".$result->{$staticOids{totalMemory}} / (1024*1024)." MB\n";
print "# of CPUs:   $result->{$staticOids{cpuCount}}\n";
print "# of blades: $result->{$staticOids{bladeCount}}\n";
print "LTM Version: $result->{$staticOids{ltmVersion}}\n";
print "LTM Build:   $result->{$staticOids{ltmBuild}}\n\n\n";

##
## Begin Main
##

# loop until start-of-test is detected
&detect_test(\%dataOids) unless $BYPASS;
if ($pause && !$BYPASS) {
  print "Pausing for ".$pause." seconds while for ramp-up\n";
  sleep($pause);
}

# start active polling
$pollTimer{'testStart'} = Time::HiRes::time();

while ($elapsed <= $testLen) {
  $pollTimer{'activeStart'} = Time::HiRes::time();


  # Query the LTM for the snmp data and return a hash containing the 
  # oid text names and the oid value
  $dataVals = &get_snmp_data(\%dataOids, \%snmpOpts);

  die("Error retrieving snmp data! (\$dataVals not defined)") if (!defined($dataVals));
  $pollTimer{'pollTime'} = Time::HiRes::time();

  # update $runTime now so it will be accurate when written to the file
  $runTime    = sprintf("%.7f", ($pollTimer{'pollTime'} - $pollTimer{'testStart'}));
  
  # Get the exact time since the previous loop ended and divide the amount
  # by which the SNMP counters increased by this value.
  if ($pollTimer{activeStop}) {
    $loopTime   = $pollTimer{pollTime} - $pollTimer{activeStop};
  }


  # Before any real processing, remove any non-numeric values (i.e. 'noSuchInstance')
  for my $n (keys(%dataOids)) {
    if ($dataVals->{$dataOids{$n}} =~ /\D+/) {
      $xlsData{$n} = 0;
    }
    else {
      $xlsData{$n} = $dataVals->{$dataOids{$n}};
    }
  }

  # calculate accurate values from counters and record results
  ($cpuUtil, $cpuPercent) = &get_cpu_usage($dataVals, \%oldData);
  ($tmmUtil, $tmmPercent) = &get_tmm_usage($dataVals, \%oldData);

  $memUsed    = $dataVals->{$dataOids{tmmTotalMemoryUsed}};
  $hMem       = sprintf("%d", $memUsed / MB);

  # client and server current connections
  $clientCurConns = $dataVals->{$dataOids{sysStatClientCurConns}};
  $serverCurConns = $dataVals->{$dataOids{sysStatServerCurConns}};

  # If requested, write the output file.
  if ($DATAOUT) {
    $row++;
    $raw_data->write($row, 0, $runTime, $formats{decimal4});
    $raw_data->write($row, 1, $cpuUtil, $formats{decimal2});
    $raw_data->write($row, 2, $tmmUtil, $formats{decimal2});

    $raw_data->write( $row, 
                      3,
                      [$memUsed,
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatClientBytesIn}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatClientBytesOut}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatClientPktsIn}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatClientPktsOut}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatServerBytesIn}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatServerBytesOut}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatServerPktsIn}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatServerPktsOut}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatClientCurConns}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatClientTotConns}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatServerCurConns}}),
                      sprintf("%.0f", $dataVals->{$dataOids{sysStatServerTotConns}})],
                     $formats{'standard'});
  }

  if ($elapsed > 0) {
    # pre-format some vars
    $cBytesIn   = sprintf("%.0f", ($dataVals->{$dataOids{sysStatClientBytesIn}}  
                                              - $oldData{cBytesIn}) / $loopTime);
    $cBytesOut  = sprintf("%.0f", ($dataVals->{$dataOids{sysStatClientBytesOut}}
                                              - $oldData{cBytesOut}) / $loopTime);
    $sBytesIn   = sprintf("%.0f", ($dataVals->{$dataOids{sysStatServerBytesIn}} 
                                              - $oldData{sBytesIn}) / $loopTime);
    $sBytesOut  = sprintf("%.0f", ($dataVals->{$dataOids{sysStatServerBytesOut}}
                                              - $oldData{sBytesOut}) / $loopTime);
    $cNewConns  = sprintf("%.0f", ($dataVals->{$dataOids{sysStatClientTotConns}}
                                              - $oldData{cTotConns}) / $loopTime);
    $sNewConns  = sprintf("%.0f", ($dataVals->{$dataOids{sysStatServerTotConns}}
                                              - $oldData{sTotConns}) / $loopTime);
    $cPktsIn    = sprintf("%.0f", ($dataVals->{$dataOids{sysStatClientPktsIn}}  
                                              - $oldData{cPktsIn}) / $loopTime);
    $cPktsOut   = sprintf("%.0f", ($dataVals->{$dataOids{sysStatClientPktsOut}} 
                                              - $oldData{cPktsOut}) / $loopTime);
    $cBitsIn    = sprintf("%.0f", (($cBytesIn * 8)  / 1000000));
    $cBitsOut   = sprintf("%.0f", (($cBytesOut * 8) / 1000000));

# Print updates to the screen during the test
    format STDOUT_TOP =
 @>>>>  @>>>   @>>>    @>>>>>>>>     @>>>>>     @>>>>>  @>>>>>>>>>  @>>>>>>>>>  @>>>>>>  @>>>>>>>
"Time", "CPU", "TMM", "Mem (MB)", "C-CPS", "S-CPS", "Client CC", "Server CC", "In/Mbs", "Out/Mbs"
.

    format =
@>>>>> @##.## @##.##    @#######  @########  @########  @#########  @#########   @#####    @#####
$elapsed, $cpuUtil, $tmmUtil, $hMem, $cNewConns, $sNewConns, $clientCurConns, $serverCurConns, $cBitsIn, $cBitsOut
.
     write;

#    format STDOUT_TOP =
#@>>>>   @>>>  @>>>>  @>>>>>> @>>>>>> @>>>>>>>>  @>>>>>>>>  @>>>>>>>> @>>>>>>>>
#"Time", "CPU", "TMM", "C-CPS", "S-CPS", "In/Mbps", "Out/Mbps", "Pkts/In", "Pkts/Out"
#.
#
#    format =
#@#### @##.## @##.## @>>>>>> @>>>>>> @>>>>>>>> @>>>>>>>> @>>>>>>>> @>>>>>>>>
#$elapsed, $cpuUtil, $tmmUtil, $cNewConns, $sNewConns, $cBitsIn, $cBitsOut, $cPktsIn, $cPktsOut
#.
#    write;
  }


  # update 'old' data with the current values to calculate delta next cycle
  $oldData{ssCpuRawUser}   = $dataVals->{$dataOids{ssCpuRawUser}};
  $oldData{ssCpuRawNice}   = $dataVals->{$dataOids{ssCpuRawNice}};
  $oldData{ssCpuRawSystem} = $dataVals->{$dataOids{ssCpuRawSystem}};
  $oldData{ssCpuRawIdle}   = $dataVals->{$dataOids{ssCpuRawIdle}};
  $oldData{tmmTotalCycles} = $dataVals->{$dataOids{tmmTotalCycles}};
  $oldData{tmmIdleCycles}  = $dataVals->{$dataOids{tmmIdleCycles}};
  $oldData{tmmSleepCycles} = $dataVals->{$dataOids{tmmSleepCycles}};
  $oldData{cBytesIn}       = $dataVals->{$dataOids{sysStatClientBytesIn}};
  $oldData{cBytesOut}      = $dataVals->{$dataOids{sysStatClientBytesOut}};
  $oldData{sBytesIn}       = $dataVals->{$dataOids{sysStatServerBytesIn}};
  $oldData{sBytesOut}      = $dataVals->{$dataOids{sysStatServerBytesOut}};
  $oldData{cPktsIn}        = $dataVals->{$dataOids{sysStatClientPktsIn}};
  $oldData{cPktsOut}       = $dataVals->{$dataOids{sysStatClientPktsOut}};
  $oldData{cTotConns}      = $dataVals->{$dataOids{sysStatClientTotConns}};
  $oldData{sTotConns}      = $dataVals->{$dataOids{sysStatServerTotConns}};


  $elapsed += $loopTime;
  #print "\$elapsed: ".$elapsed."\n";

  # Calculate how much time this polling cycle has required to determine how
  # long we should sleep before beginning the next cycle
  $pollTimer{activeStop}    = Time::HiRes::time();
  $pollTimer{activeRunTime} = $pollTimer{activeStop} - $pollTimer{activeStart};
  $sleepTime = $cycleTime - $pollTimer{activeRunTime};

  $loopTotal = $pollTimer{'activeStop'} - $lastLoopEnd;


### uncomment for debugging purposes only
#format STDOUT_TOP =
#@>>>>>>>>  @||||||||||||  @||||||||||||||||| @||||||||||||||||| @||||||||||||||||| @||||||||||||||||| @|||||||||  @|||||||||  @|||||||||
#"RunTime", "Elapsed", "activeStart", "activeStop", "pollTime", "activeRunTime", "loopTime", "sleepTime", "lastLoopTotal"
#.
#
#  format =
#@########  @######.#####  @###########.##### @###########.##### @###########.#####     @#.#####       @##.#####   @##.#####   @##.#####
#$runTime, $elapsed, $pollTimer{activeStart}, $pollTimer{activeStop}, $pollTimer{pollTime}, $pollTimer{activeRunTime}, $loopTime, $sleepTime, $loopTotal
#.
#    write;
  
  $lastLoopEnd = Time::HiRes::time();
  Time::HiRes::sleep($sleepTime);
} 



##
## Subs
##

# delay the start of the script until the test is detected through pkts/sec
sub detect_test() {
  my $oids = shift;
  my $pkts = 3000;

  print "\nWaiting for test to begin...";

  while (1) {
    my $r1 = `snmpget -t2 -Ovq -c public localhost $$oids{sysStatClientPktsIn}`;
    sleep(2);
    my $r2 = `snmpget -t2 -Ovq -c public localhost $$oids{sysStatClientPktsIn}`;

    my $delta = $r2 - $r1;
  
    if ($delta > $pkts) {
      print "\nStart of test detected...\n\n";
      return;
    }
    else {
      print ".";
      sleep(3);
    }
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
      'sysStatClientBytesIn'    => '.1.3.6.1.4.1.3375.2.1.1.2.1.3.0',
      'sysStatClientBytesOut'   => '.1.3.6.1.4.1.3375.2.1.1.2.1.5.0',
      'sysStatClientPktsIn'     => '.1.3.6.1.4.1.3375.2.1.1.2.1.2.0',
      'sysStatClientPktsOut'    => '.1.3.6.1.4.1.3375.2.1.1.2.1.4.0',
      'sysStatClientTotConns'   => '.1.3.6.1.4.1.3375.2.1.1.2.1.7.0',
      'sysStatClientCurConns'   => '.1.3.6.1.4.1.3375.2.1.1.2.1.8.0',
      'sysStatServerBytesIn'    => '.1.3.6.1.4.1.3375.2.1.1.2.1.10.0',
      'sysStatServerBytesOut'   => '.1.3.6.1.4.1.3375.2.1.1.2.1.12.0',
      'sysStatServerPktsIn'     => '.1.3.6.1.4.1.3375.2.1.1.2.1.9.0',
      'sysStatServerPktsOut'    => '.1.3.6.1.4.1.3375.2.1.1.2.1.11.0',
      'sysStatServerTotConns'   => '.1.3.6.1.4.1.3375.2.1.1.2.1.14.0',
      'sysStatServerCurConns'   => '.1.3.6.1.4.1.3375.2.1.1.2.1.15.0',
      'sysIpStatDropped'        => '.1.3.6.1.4.1.3375.2.1.1.2.7.4.0',
                );

  return(%oidlist);
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
  );

  return(%oidlist);
}


# returns a hash containing the oid and data
sub get_snmp_data() {
  my $oids = shift;
  my $opts = shift;

  my $oidlist = "";
  my @output;
  my ($oid, $value);
  my (%results, %data);

  for my $o (values(%$oids)) { $oidlist = $oidlist.' '.$o; };
  @output = `snmpget -t2 -Onq -c $opts->{comm} $opts->{host} $oidlist`;
  
  # parse each line of the response for the oid and the value
  foreach (@output) {
    chomp($_);
    my ($oid, $value) = split(' ', $_);
    $data{$oid} = $value;
  }
  # return a hash with the text oid names and the data values
  return(\%data);
}

# returns a hash containing the oid names and data (rather than oid name and oid id)
sub get_snmp_data2() {
  my $oids = shift;
  my $opts = shift;

  my $oidlist = "";
  my @output;
  my ($oid, $value);
  my (%results, %data);

  for my $o (values(%$oids)) { $oidlist = $oidlist.' '.$o; };
  @output = `snmpget -t2 -Onq -c $opts->{comm} $opts->{host} $oidlist`;
  
  # parse each line of the response for the oid and the value
  foreach (@output) {
    chomp($_);
    my ($oid, $value) = split(' ', $_);
    # match the oid in the response to the oid in %$oids
    while (my ($oidName, $oidID) = each(%$oids)) {
      if ($oid =~ $oidID) {
        # populate %data with the oid text name and the data value
        $data{$oidName} = $value;
      }
    }
  }
  # return a hash with the text oid names and the data values
  return(\%data);
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


# print script usage and exit with the supplied status
sub usage() {
  my $code = shift;

  print <<END;
  USAGE:  $0 -d <host> -l <total test length> -o <output file>
          $0 -h

  -d      IP or hostname to query (REQUIRED)
  -l      Full Test duration             (default: 130 seconds)
  -i      Seconds between polling cycles (default: 10 seconds)
  -o      Output filename                (default: /dev/null) - NOT IMPLEMENTED
  -p      Pause time; the amount of time to wait following the start of the 
          test before beginning polling. Should match the ramp-up time (default: 0)
  -v      Verbose output (print stats)
  -B      Bypass start-of-test detection and start polling immediately
  -h      Print usage and exit

END

  exit($code);
}
