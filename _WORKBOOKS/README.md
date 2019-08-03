<!--
  Author: Cesar Raul Urteaga-Reyesvera.
-->

# Forms

## Table of Contents

-   [FShowHideSheets.xlsm](#fshowhidesheets)
-   [MGetHierarchicalTreeContents.xlsb](#mgethierarchicaltreecontents)

### FShowHideSheets

This workbook shows how to use the FShowHideSheets form that renders information as to the sheets of the active workbook as it is depicted below:

<img src="_IMAGES\FShowHideSheets.png">

In case you only want the userform please check out <a href="https://github.com/Cesar-Urteaga/1907_EXCEL/tree/master/_FORMS" target="_blank">https://github.com/Cesar-Urteaga/1907_EXCEL/tree/master/_FORMS</a>.

### MGetHierarchicalTreeContents

This workbook shows how to use the <a href="https://github.com/Cesar-Urteaga/1907_EXCEL/blob/master/_CODES/MGetHierarchicalTreeContents.bas" target="_blank">MGetHierarchicalTreeContents</a> code that recursively gets the information of the contents of a folder and displays them in a table as it is shown hereunder:

* First, the macro prompts the user to select the folder for which its contents are required.
<img src="_IMAGES\MGetHierarchicalTreeContents_01.png">

* Second, once the user specifies the folder (i.e., the host folder), it creates a table with all its contents.
<img src="_IMAGES\MGetHierarchicalTreeContents_02.png">

The definitions of the non-self-explanatory fields are the following:

 * Type         : Type of the object (i.e., folder [D] or file [F]).
 * Name         : Name of the object.
 * Folder Path  : Parent folder where the object is stored (displayed as a hyperlink).
 * File Path    : If it is a file, it is its path; otherwise, it holds the same value as "Folder Path". Also, it is displayed as a hyperlink.
 * \#           : Number level of hierarchy.
 * Hierarchy    : Graphically displays the hierarchy ("|" for a folder and "\*" for a file).
