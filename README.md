# utl_side_by_side_excel_reports
Create Side by Side excel reports starting at arbitrary positions

    ```    ```
    ```    ```
    ```  T1006790 SAS Forum: SAS/WPS/R Two reports one starting at B3 and the second at G3 (side by side)  ```
    ```    ```
    ```    ```
    ```     WORKING CODE  ```
    ```    ```
    ```          writeWorksheet(wb,   males, sheet = "sex", startRow = 3,startCol = 2, header = TRUE);  ```
    ```          writeWorksheet(wb, females, sheet = "sex", startRow = 3,startCol = 7, header = TRUE);  ```
    ```    ```
    ```  see  ```
    ```  https://goo.gl/4VfdMR  ```
    ```  https://communities.sas.com/t5/ODS-and-Base-Reporting/How-to-print-report-in-a-specific-cell-in-excel/m-p/378853  ```
    ```    ```
    ```    ```
    ```  I am using SAS 9.3 version and I am trying to output some data to excel  ```
    ```  starting from a particular row and column number. In the attached excel  ```
    ```  screenshot I have two reports. The first one is starting at Column B Row  ```
    ```  3 and the second report is starting at Column G Row 3 and so forth.  ```
    ```    ```
    ```  Could someone show how to direct SAS to start printing the report from a  ```
    ```  particular row and column cell in excel. Thanks for your help!  ```
    ```    ```
    ```    ```
    ```  HAVE  ```
    ```  ====  ```
    ```    ```
    ```    Up to 40 obs SD1.HAVE total obs=10  ```
    ```    ```
    ```    Obs    NAME       SEX  ```
    ```    ```
    ```      1    Alfred      M  ```
    ```      2    Alice       F  ```
    ```      3    Barbara     F  ```
    ```      4    Carol       F  ```
    ```      5    Henry       M  ```
    ```      6    James       M  ```
    ```      7    Jane        F  ```
    ```      8    Janet       F  ```
    ```      9    Jeffrey     M  ```
    ```     10    John        M  ```
    ```    ```
    ```    ```
    ```    ```
    ```    ```
    ```   WANT excel sheet with males starting at A3 and females at G3  ```
    ```  ==============================================================  ```
    ```    ```
    ```  d:/xls/sex_fm.xlsx  ```
    ```    ```
    ```      +---------------------+-----------------------------------+  ```
    ```      |  A  |  B    |  C    |  D  |  E  |  F    |  G    |  H    |  ```
    ```      +---------------------+-----------------------------------+  ```
    ```  1   |     |       |       |     |     |       |       |       |  ```
    ```      |-----+-------+-------|-----+-----+-------+-------+-------|  ```
    ```  2   |     |       |       |     |     |       |       |       |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  3   |     |NAME   |SEX    |     |     |       |NAME   |SEX    |  ```
    ```      |-----+-------+-------|-----+-----+-------+-------+-------|  ```
    ```  4   |     |Alfred |M      |     |     |       |Alice  |M      |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  5   |     |Alex   |M      |     |     |       |Barbara|M      |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  6   |     |Bob    |M      |     |     |       |Carol  |M      |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  7   |     |Chris  |M      |     |     |       |Jane   |M      |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  8   |     |Henry  |M      |     |     |       |       |       |  ```
    ```      -----------------------------------------------------------  ```
    ```  ...  ```
    ```    ```
    ```  [SEX]  ```
