Files

ltmStats-dev.pl
  Current development version branched from ltmStats-json.pl

ltmStats.pl
  Remote statistics collection and recording utility. The script collects
  performance data using SNMP and writes the collected statistics to an Excel
  spreadsheet (xlsx) or a json file (under development).

dutStats.pl
  A reduced-functionality version of ltmStats.pl. This version does run
  on LTM and print real-time performance data to the screen just like
  ltmStats.pl, but it does not write an output file, so the results
  are not recorded.  This is the version of the script that we typically
  make available to customers or SEs (upon request).

ltmStats-vipStats.pl
  An SE-contributed version of ltmStats.pl that writes per-virtual-server 
  data to the spreadsheet. The version of ltmStats.pl that this is based
  on pre-dates several recent updates, and as a result does not 
  incorporate the latest options in ltmStats. The per-VS modifications 
  the SE contributed (Grzegorz Kornacki) should be imported into ltmStats.pl.
