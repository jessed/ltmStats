Files

ltmStats.pl
  Remote statistics collection and recording utility. The script collects
  performance data using SNMP and writes an Excel spreadsheet (xlsx) 
  using the Excel::Spreadsheet::WriteExcel Perl module. This version
  does not run directly on LTM due to the lack of the Excel::
  Spreadsheet::WriteExcel module.

dutStats.pl
  A reduced-functionality version of ltmStats.pl. This version does run
  on LTM and print real-time performance data to the screen just like
  ltmStats.pl, but it does not write an output file, so the results
  are not recorded.
  This is the version of the script that we typically make available 
  to customers or SEs (upon request).

archive.tar.gz
  Old versions of ltmStats/dutStats prior to the project being
  added to version control. Also contains some one-off scripts
  related to stats collection, and utility scripts written
  when ltmStats was first being developed.
  This is being kept only for historical context. None of the files
  in this directory are current tools; all have been superceded
  by more recent versions of ltmStats or dutStats.

