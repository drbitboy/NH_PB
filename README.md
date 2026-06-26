# New Horizons Spacecraft conversion of MET (Mission Elapsed Time) to UTC

## for eXcel and LibreOffice Calc ("Calc") workbooks

## Usage

* Copy the **sclk_data** worksheet from **met2utc.xlsm** or **met2utc.ods** as a new worksheet into a workbook
* Either import the script **excel_met2utc.bas** into that same eXcel workbook (.xlsm),
* Or import the Macro **librecalc_met2utc.bas** into that same Calc workbook.
  * And enable macros in that Calc workbook
* Use the expression **=met2utc(met)** anywhere in that workbook to convert MET to UTC

## Caveats and notes

* The MET argument is a numeric value, equal to approximately (year-2006) x $\pi x 10^7$
* The MET argument will typically be a Cell reference e.g. **met2utc(A3)**, but it could also be an expression
* The MET-to-UTC piecewise linear interpolation data have a limited range; as of 26-June-2026:
  * The Spacecraft CLock Kernel (NAIF/SPICE SCLK) used is **new_horizons_3381.tsc**
    * The last entry in that SCLK is ~August, 2025
    * The Python script **met2utc.py** assumes the time column in the SCLK is TDT
      * Kernel Pool Value SCLK01_TIME_SYSTEM_98 is 2
  * The LeapSecond Kernel (LSK) is **naif0012.tls**
    * The last accumulated leapseconds value (37) in that LSK is for 01.January, 2017
    * The offset between TDT and UTC is 69.184s (37 + 32.184 = TDT - UTC)
  * The start of the range is ~January, 2017, just after the most recent (37th) leapsecond
  * The end of the range is ~August, 2045, 20y past the last SCLK entry

## Source files

* met2utc.xlsm - eXcel sample workbook
* met2utc.ods - LibreOffice Calc sample workbook
* excel_met2utc.bas - eXcel VBA code file with function met2utc
* librecalc_met2utc.bas - Calc Macro Basic file with function met2utc
* met2utc.csv - data for worksheet sclk_data as Comma-Separated Values
* met2utc.py - generate met2utc.csv from SPICE kernels
* naif0012.tls, new-horizons_3381.tsc - default SPICE kernels
* mk.tm - utility meta-kernel, not other than for development

## Background

* Uses NAIF/SPICE Spacecraft CLock and LeapSecond kernels (SCLK and LSK) to set up the piecewise interpolation
* Python script **met2utc.py** uses module **spiceypy**
  * Writes a CSV to stdout e.g. ```python met2utc.py > metutc.py```
    * Those CSV data that can be copied as-is into worksheet **sclk_data**
    * Cell range **A1:G*nnn***, where ***nnnn*** is the last row
* Google AI mode links for some of this code
  * (https://share.google/aimode/a1MbYq1lSD0cOYC5P)[https://share.google/aimode/a1MbYq1lSD0cOYC5P]
  * (https://share.google/aimode/X1DAbrQV7ZQ0xWlma)[https://share.google/aimode/X1DAbrQV7ZQ0xWlma]
  * (https://share.google/aimode/yOBTuCytr2EVJHPjq)[https://share.google/aimode/yOBTuCytr2EVJHPjq]
  * N.B.
    * For the eXcel VBA and Libre Calc Macro Basic scripts
    * These links may expire
* NAIF/SPICE Toolkit references
  * [Time Required Reading](https://naif.jpl.nasa.gov/pub/naif/toolkit_docs/C/req/time.html)
  * [SCLK Required Reading](https://naif.jpl.nasa.gov/pub/naif/toolkit_docs/C/req/sclk.html)
