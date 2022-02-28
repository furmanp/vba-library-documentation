# FAQ

## How to install library on my computer?
To install the library go to <a href="https://github.com/furmanp/Pontis-VBA" target="_blank" rel="noopener">GitHub repository</a> and clone it to your local machine.
You can either manually download the .zip file or if you have GitHub installed you can clone the repository via the app. This will allow you to update the project once new
changes are introduced. 
Next, go to the MS Excel spreadsheet and import downloaded modules. Save the workbook as **`PERSONAL.XLSB`** under **`C:/Users/User Name/AppData/Microsoft/Excel/XLSTART`**.
This will provide access to your newly imported functions for all spreadsheets in your computer. 
To activate the library, open VBE in MS Excel and go to **`Tools/References`** and select VBAPROJECT unless you named it otherwise.


## How to update the documentation?
To contribute to the documentation and add description of newly developed functions follow steps below.

!!! note
    for seamless process of updating the documentation you should have programming IDE & GitHub installed on your machine.

* clone the <a href="https://github.com/furmanp/vba-library-documentation" target="_blank" rel="noopener">Library repository</a> on your local machine,
* navigate to **`docs/pages/docs`** and open file you want to update,
* copy table template from the **`vba_documentation_template.md`** in the same catalogue,
* fill in the table with detailed description to make the function clear and easily understandable to others,
* once finished, commit your changes to the repository and create a pull request with descriptive message,
* assign reviewer to check the code and merge changes.

