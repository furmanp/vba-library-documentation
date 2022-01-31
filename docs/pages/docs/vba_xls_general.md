# Excel/General

### *Enum vOutput*
Enumeration type variable listing possible outputs when handling functions that work with worksheet cells. <br>
<ul>
 <li> vAddress <br></li>
 <li> vValue <br></li>
 <li> vWidth <br></li>
 <li> vLength <br></li>
</ul>

### *xls_import_data*
|xls_import_data <br> (ws_main, s_keyword)|
|:----------------------------------------|
|Description:|
|Method to import data from chosen .xlsx file to main workbook. Data will be imported to sheet that is active while the method is being called. Cell *G2* will have a  full directory path of the imported file pasted.|
|Input:|
|**STRING `ws_main`** &emsp; Name of the worksheet, in the Workbook the data is being imported from.|
|**STRING `s_keyword`** &nbsp; Keyword that should indicate top right corner of the data that is to be imported.|
|Output:|
|None|
|Return Code:|
|None|
|Remarks/Usage:|
|It is recommended to use format with preceding dot and a number at the end *.data1* to avoid confusion with regularly used expressions. Method supports import of multiple tables, as long as they match the keyword *.data, .data1, .data2, .dataN*. First occurrence of the keyword does not need to have a number. <br> Method import_data is calling functions **`xls_look`**, **`a_make`**, **`a_paste`**, **`a_tidy`**, and **`WorksheetExists`** from the same module.|
|Example:|
|None|


### *a_base0*
|a_base0 <br> (input_arr)|
|:-----------------------| 
|Description:|
|Function takes in array indexed from 1 and updates indexing to start from 0. Supports 1 dimensional arrays only.|
|Input:|
|**VARIANT `input_arr`** &nbsp; Option Base 1 array to be updated to Option Base 0.|
|Output:|
|**VARIANT `a_base0`** &emsp;&nbsp; Array with indexing starting from 0. No additional output arguments required.|
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *a_test*
|a_test <br> (input_arr)|
|:----------------------|
|Description:|
|Subroutine pastes passed array into worksheet called *test*. If the workbook exists, it replaces current content of the worksheet. If not, routine creates one.|
|Input:|
|**VARIANT `input_arr`** &nbsp; Array that is pasted into cell B2, in *test* Worksheet.|
|Output:|
|None|
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *a_paste*
|a_paste <br> (input_arr, paste_cell, offset_row, offset_col, wbook)|
|:------------------------------------------------------------------|
|Description:|
|Soubroutine that pastes passed array into passed cell.|
|Input:|
|**VARIANT `input_arr`** Array that subroutine will paste into defined cells.|
|**STRING `paste_cell`** Destination cell where the data is going to be inserted to. Can be provided in format of absolute address (for example *A1*), or a keyword (*.data*).|
|**INTEGER `[offset_row, offset_col]`** Values that define whether the data should be inserted in a found cell (default option), or offset. Default value is 0 (direct location).|
|**WORKBOOK `[wbook]`** Allows to look for data in different workbooks to the active one. It will look for in active sheet though. Default value is ThisWorkbook.|
|Output:|
|None|
|Return Code:|
|None|
|Remarks/Usage:|
|Subroutine is calling in **`xls_look`** function.|
|Example:|
|None|

### *a_look*
|a_look <br> (s_keyword, input_arr, num)|
|:--------------------------------------|
|Description:|
|Function looks for **`s_keyword`** in first row and column of **`input_arr`** and returns value defined by **`num`**.|
|Input:|
|**VARIANT `s_keyword`** &nbsp; Keyword that is to be found in row/column header.|
|**VARIANT `input_arr`** &nbsp; Array that the **`s_keyword`** is being looked in.
|**INTEGER `num`** &emsp;&emsp;&emsp;&ensp;Values [-1, 0, 1, 2, 3... n < Ubound(input_arr)]
|Output:|
|**VARIANT `a_paste`**  &emsp;&ensp;Function returns accordingly:<br> <ul><li>-1; column/row index as INTEGER<br></li><li> 0; column/row index with preceding name (col/row) to indicate s_keyword location.<br></li><li> 1, 2, 3... n - value that is stored in cell 1, 2, 3... cells below/right to the s_keyword cell.</li></ul>|
|Return Code:|
|None|
|Remarks/Usage:|
|Function is calling in a **`a_2Dto1D`** function.
|Example:|
|None|

### *a_make*
|a_make <br> (s_keyword, n_down_start, n_right_start, n_down_end, n_right_end, wbook)|
|:------------------------------------------------------------------------------|
|Description| 
|Function sweeps active worksheet for s_keyword value.|
|Input:|
|**STRING `s_keyword`** Keyword that is to be found in row/column header.|
|**LONG `[n_down_start, n_right_start]`** Parameters define whether function should start from the s_keyword cell of cell adjacent offset down/right by passed value. Default value equals 0.|
|**LONG `[n_down_end, n_right_end]`** Parameters define length/width of created matrix. If omitted, function will take entire table which represents combination of keys ++ctrl+shift+down++ |
|Output:|
|**VARIANT `a_make`** 2 dimensional matrix of data that was stored in continuous chain of cells to the right and to the bottom of s_keyword cell, unless offsets were specified.|
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *a_tidy*
|a_tidy <br>()|
|:---------------------------------------|
|Description:|
||
|Input:|
||
|Output:|
||
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *isVariantAllocated*
|isVariantAllocated <br>()|
|:---------------------------------------|
|Description:|
||
|Input:|
||
|Output:|
||
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *xls_WorksheetExists*
|xls_WorksheetExists <br>(ws_name, wbook)|
|:---------------------------------------|
|Description:|
|Function checks if worksheet with passed ws_name in declared workbook exists. If no wbook is declared, function will look in workbook that is calling the funcion.|
|Input:|
|**STRING `ws_name`** &emsp;&emsp;&ensp; Name of the worksheet to be checked if exists.|
|**WORKBOOK `[wbook]`** &nbsp; Workbook where the function should look for ws_name. Default value is `ThisWorkbook`.|
|Output:|
|**BOOLEAN `xls_WorksheetExists`** &nbsp; Boolean statement. Returns `True` is Worksheet named ws_name exists, returns `False` otherwise.
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *xls_fast*
|xls_fast <br>(status)|
|:--------------------|
|Description:|
|Subroutine to slightly improve Worksheets related macro efficiency. It turns off Screen Updating and calculation. It is recommended to set is back on at the end of the script.|
|Input:|
|**BOOLEAN `[status]`** Variable to set the function on or off. If omitted, default value is True.|
|Output:|
|None|
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *array_1Dto2D*
|array_1Dto2D <br>(input_arr)|
|:---------------------------|
|Description:|
|Function converts 1 dimensional array and allocates second dimension for each item. Returns 2 dimensional array.|
|Input:|
|**VARIANT `input_arr`** &nbsp; 1 dimensional array that will have second dimension allocated.|
|Output:|
|**VARIANT `array_1Dto2D`** 2 dimensional array. Second dimension is 1 element vector.|
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *xls_inputbox*
|xls_inputbox <br>()|
|:------------------|
|Description:|
|Function displays an InputBox asking for user input. User can provide either sinlge values separated with a comma, or provide limits of a range separated with a dash. Inserted data is then converted into an array.|
|Input:|
|No input variables required. User input is passed through an InputBox textfield.|
|Output:|
|**VARIANT `xls_inputbox`** &nbsp; 2 dimensional vector of values read from input box.|
|Return Code:|
|None|
|Remarks/Usage:|
|<ul><li>Input: 1, 2, 3, 4, 5 -> Output: array = (1, 2, 3, 4, 5) <br></li><li> Input: 1- 10 -> Output: array = (1, 2, 3, 4... , 10) <br></li></ul>Negative ranges are supported.|
|Example:|
|None|

### *xls_look*
|xls_look <br>(s_keyword, output, offset_down, offset_right, wbook)|
|:---------------------------------------------------------------------|
|Description:|
|Function sweeps active worksheet for s_keyword value. Returns data dependent on the input choice.|
|Input:|
|**STRING `s_keyword`** &nbsp; Keyword that the function is looking for in the active worksheet.|
|**vOutput `output`** &emsp;&emsp;Enumeration variable indicating what function should return. {**`vValue`**, **`vAddress`**, **`vLength`**, **`vWidth`**}
|**LONG `[offset_down, offset_right]`** Indicates whether function should return values offset from cell containing s_keyword. If omitted, function will return information about cell containing **`s_keyword`** (exact match).|
|**WORKBOOK `[wbook]`** Allows to look for data in different workbooks to the active one. It will look for the value in active sheet though.|
|Output:|
|**VARIANT `xls_look`** &emsp;Returned value depends on value chosen in output variable.|
|Return Code:|
|None|
|Remarks/Usage:|
|None|
|Example:|
|None|

### *a_nDim*
|a_nDim <br>(input_arr)|
|:---------------------|
|Description:|
|Function returns number of dimensions of input_arr that was passed on to the function. Returns a Long variable informing about number of dimensions.|
|Input:|
|**VARIANT `input_arr`** &nbsp; Array that number of dimension will be calculated from.|
|Output:|
|**LONG `a_nDim`** &emsp;&emsp;&emsp;&nbsp;Value that represents number of dimensions array has.|
|Return Code:|
|None|
|Remarks/Usage|
|None|
|Example:|
|None|

### *a_2Dto1D*
|a_2Dto1D <br>(input_arr, a_vector)|
|:---------------------------------|
|Description:|
|Function extracts single dimensional vector from 2 dimensional matrix.|
|Input:|
|**VARIANT `input_arr`** &nbsp;2 dimensional matrix that row or column is to be extracted from.|
|**STRING `a_vector`** &emsp;&nbsp; Keyword that indicated which row/column should be extracted from input_arr matrix. Syntax should look as follows: row1, col3.
|Output:|
|**VARIANT `a_2Dto1D`** &emsp; Single dimensional vector.|
|Return Code:|
|None|
|Remarks/Usage:|
|**`a_vector`** should indicate whether user wants to extract row or column and point at it's index (base1).<br> For example: row1, col3.|
|Example:|
|None|

### *a_len*
|a_len <br>(input_arr)|
|:--------------------|
|Description:|
|Function returns lenght of the array as **Long**|
|Input:|
|**VARIANT `input_arr`** &nbsp; Array that length will be returned|
|Output:|
|**STRING `a_len`** Length of the array (first dimension only).|
|Return Code:|
|None|
|Remarks/Usage:|
|Indexing the array (base0, base1) does not affect the result. To calculate length of array's second dimension rin **`a_len(arr[1])`**.| 
Example:|
|None|

