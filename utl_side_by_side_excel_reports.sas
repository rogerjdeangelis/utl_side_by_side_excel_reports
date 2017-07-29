
*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;
options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.have;
  set sashelp.class(keep=name sex obs=10);
run;quit;

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|
;

%utl_submit_r64('
source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);
library(haven);
library(XLConnect);
have<-read_sas("d:/sd1/have.sas7bdat");
have;
males<-have[have$SEX=="M",];
females<-have[have$SEX=="F",];
wb <- loadWorkbook("d:/xls/sex_mf.xlsx", create = TRUE);
createSheet(wb, name = "sex");
writeWorksheet(wb, males, sheet = "sex", startRow = 3,startCol = 2, header = TRUE);
writeWorksheet(wb, females, sheet = "sex", startRow = 3,startCol = 7, header = TRUE);
saveWorkbook(wb);
');

