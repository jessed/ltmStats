LTM statistics collection via SNMP. Outputs data to CLI in realtime and to xlsx (if appropriate modules are present).

Files

devStats
  Development version

ltmStats
  Main/stable version

summarize_tests.py
  Scans the current directory for ltmStats output (*xlsx), reads all the files and 
  creates summarized output xlsx.


deprecated/*
  Older scripts no longer intended for use. Maintained here only to allow slightly faster
  recovery if a problem is found with the stable version in the middle of an engagement.

ltmStats-vipStats.pl
  An SE-contributed version of ltmStats.pl that writes per-virtual-server 
  data to the spreadsheet. The version of ltmStats.pl that this is based
  on pre-dates several recent updates, and as a result does not 
  incorporate the latest options in ltmStats. The per-VS modifications 
  the SE contributed (Grzegorz Kornacki) should be imported into ltmStats.pl.
