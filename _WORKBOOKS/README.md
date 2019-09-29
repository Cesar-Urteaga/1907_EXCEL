<!--
  Author: Cesar Raul Urteaga-Reyesvera.
-->

# Workbooks

## Table of Contents

-   [FShowHideSheets.xlsb](#fshowhidesheetsxlsb)
-   [MGetHierarchicalTreeContents.xlsb](#mgethierarchicaltreecontentsxlsb)
-   [MAnimate.xlsb](#manimatexlsb)

### FShowHideSheets.xlsb

This workbook shows how to use the FShowHideSheets form that renders information as to the sheets of the active workbook as it is depicted below:

<img src="_IMAGES\FShowHideSheets.png">

In case you only want the userform please check out <a href="https://github.com/Cesar-Urteaga/1907_EXCEL/tree/master/_FORMS" target="_blank">https://github.com/Cesar-Urteaga/1907_EXCEL/tree/master/_FORMS</a>.

### MGetHierarchicalTreeContents.xlsb

This workbook shows how to use the <a href="https://github.com/Cesar-Urteaga/1907_EXCEL/blob/master/_CODES/MGetHierarchicalTreeContents.bas" target="_blank">MGetHierarchicalTreeContents</a> code that recursively gets the information of the contents of a folder and displays them in a table as it is shown hereunder:

* First, If it is not specified none of the macro parameters, by default the macro will prompt the user to select the folder for which its contents are required.
<img src="_IMAGES\MGetHierarchicalTreeContents_01.png">

* Second, once the user specifies the folder (i.e., the host folder), it creates a table with all its contents, starting from the active cell (i.e., the upper-left corner).
<img src="_IMAGES\MGetHierarchicalTreeContents_02.png">

Note that the table is sorted by the `File Path` field ascendantly.  Also, an autofilter is embedded.

The definitions of the non-self-explanatory fields are the following:

  * `Type`         : Type of the object (i.e., folder [D] or file [F]).
  * `Name`         : Name of the object.
  * `Folder Path`  : Parent folder where the object is stored (displayed as a hyperlink).
  * `File Path`    : If it is a file, it is its path; otherwise, it holds the same value as `Folder Path`. Also, it is displayed as a hyperlink.
  * `#`            : Number level of hierarchy.
  * `Hierarchy`    : Graphically displays the hierarchy ("|" for a folder and "\*" for a file).

On the other hand, the `DisplayHierarchicalContent` macro has the succeeding 3 optional parameters:

  * `rngPivotCell`    : Cell that indicates the upper-left corner of the table that will have the contents.  If it is not specified, it will be the active cell.
  * `sHostFolderPath` : Path of the host folder for which the information will be extracted.  If it is missing, it will prompt the user to pick out a host folder.
  * `sStartingFolder` : String that states the starting path in which the pop-up window that requests the host folder will start off.

### MAnimate.xlsb

This workbook shows how to make an emoticon animation within a cell:

<img src="_IMAGES\MAnimate01.png">
<img src="_IMAGES\MAnimate02.png">

The VBA code is provided in the <a href="https://github.com/Cesar-Urteaga/1907_EXCEL/blob/master/_CODES/MAnimation.bas" target="_blank">MAnimation.bas</a> file.
