LTM statistics collection via SNMP. Outputs data to CLI in realtime and to xlsx (if appropriate modules are present).

# Usage 
./ltmStats -h
  USAGE:  ./ltmStats -d <host> -l <total test length> -o <output file>
          ./ltmStats -h

  -d      IP or hostname to query (REQUIRED)   
  -p      UDP port to connect use         (default: 161)   
  -c      SNMP community string           (default: public)   
  -s      SNMP version                    (default: v2c)   
  -l      Full Test duration              (default: 130 seconds)   
  -i      Seconds between polling cycles  (default: 10 seconds)   
  -o      XLSX output filename            (default: /dev/null)   
  -j      JSON output filename            (default: /dev/null)   
  -P      Wait time; the amount of time to wait following the start of the   
          test before beginning polling. Should match the ramp-up time (default: 0)   
  -v      Verbose output (print verbose stats)   
  -P      Pretty output  (Stats with digital grouping)   
  -B      Bypass start-of-test detection and start polling immediately   
  -h      Print this help text and exit   

  -m      IP or hostname to monitor for the start of the test. Use this to   
          monitor the active LTM in a failover pair, but record data from the   
          standby LTM.   


# Files

## ltmStats
  Main/stable version

## other/
  summarize_tests.py  
  Scans the current directory for ltmStats output (xlsx), reads all the files and 
  creates summarized output xlsx.

  ltmStats-vipStats.pl - *Contributed by Gregorz Kornacki*   
  An SE-contributed version of ltmStats.pl that writes per-virtual-server   
  data to the spreadsheet. The version of ltmStats.pl that this is based  
  on pre-dates several recent updates, and as a result does not   
  incorporate the latest options in ltmStats. The per-VS modifications  
  the SE contributed (Grzegorz Kornacki) should be imported into ltmStats.pl.  

