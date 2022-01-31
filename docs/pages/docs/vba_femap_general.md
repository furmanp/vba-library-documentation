# Femap/General

##Mardkown styled table


| **xls_import_data**   <br> (ws_main, s_keyword)|   
| :----------------------------------------------| 
|Description:|
|Method to import data from chosen .xlsx file to main workbook. Data will be imported to sheet that is active while the method is being called. Cell “G2” will have a  full directory path of the imported file pasted.|
| Input          |
| **STRING `ws_main`** &emsp; Name of the worksheet, in the Workbook the data is being imported from.| 
| **STRING `s_keyword`** &nbsp; Keyword that should indicate top right corner of the data that is to be imported. | 
|Output:|
|None|
|Return Code:|
|None|
|Remarks/Usage:|
|It is recommended to use format with preceding dot and a number at the end .data1 to avoid confusion with regularly used expressions. Method supports import of multiple tables, as long as they match the keyword .data1, .data2, .dataN. First occurrence of the keyword does not need to have a number. Method **`import_data`** is calling functions **`xls_look`**, **`a_make`**, **`a_paste`** and **`WorksheetExists`** from the same module.|
|Example:|
|None|

##HTML styled table

<table>
    <thead>
        <tr>
            <th colspan=2 align=left>xls_import_data <br> (ws_main, s_keyword)</th>
        </tr>
    </thead>
    <tbody>
      <tr>  
        <td colspan=2>Description</td>
      </tr>
      <tr>
          <td colspan=2>
            Method to import data from chosen .xlsx file to main workbook. 
            Data will be imported to sheet that is active while the method is being called.
            Cell “G2” will have a  full directory path of the imported file pasted.
          </td>
      </tr>
      <tr>
        <td colspan=2>Input:</td>
      </tr>
      <tr>
        <td><b>STRING</b> ws_main</td>
        <td>Name of the worksheet, in the Workbook the data is being imported from.</td>
      </tr>
      <tr>
        <td><b>STRING</b> s_keyword</td>
      <td>Keyword that should indicate top right corner of the data that is to be imported.</td>
      </tr>
      <tr>
      <td colspan=2>Output:</td>
      </tr>
      <tr>
        <td>None</td>
      </tr>
      <tr>
      <td colspan=2>Return Code:</td>
      </tr>
      <tr>
        <td>None</td>
      </tr>
      <tr>
      <td colspan=2>Remarks/Usage:</td>
      </tr>
      <tr>
        <td colspan=2>It is recommended to use format with preceding dot and a number at the end .data1 to avoid confusion with regularly used expressions. 
                        Method supports import of multiple tables, as long as they match the keyword .data1, .data2, .dataN. 
                        First occurrence of the keyword does not need to have a number.
                        <br>Method import_data is calling functions xls_look, a_make, a_paste and WorksheetExists from the same module.
        </td>
      </tr>
      <tr>
        <td colspan=2>Example:</td>
      </tr>
      <tr>
        <td>None</td>
      </tr>
    </tbody>
</table>


