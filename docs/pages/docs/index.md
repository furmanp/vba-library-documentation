# Getting Started

This document consist of list of all modules and methods that have been developed throughout the time to 
help and improve efficiency of working with MS Excel, Femap, and other software. <br>
All the code provided is in Visual Basic for Application.

Contents of the file is divided into modules (which are organized in a manner of what they are mainly interacting with), next methods/functions are listed with further explanation and general use purpose.

Each section refers to the software methods primarily interact with.<br>
Methods have been grouped into categories to simplify navigation throughout the documentation. 
And so, if user is looking for a methods to manage Materials object within Femap environment go check
Femap/Materials section, and so on. <br>

All the files can be downloaded form GitHub repository which is linked in the top right corner of this website.

## How to use
###MS Excel
* Copy modules on your hard drive,
* Create new MS Excel Spreadsheet and enter Visual Basic Editor,
* Go to File/Import (Ctrl + M) and select downloaded modules,
* Use imported modules/functions.

!!! warning 
    Keep in mind that certain methods are not entirely independent and use other methods stored in this library. Documentation 
    specifically explains which dependencies are used in the method, so if you want to keep your project compact ensure you copied 
    necessary methods.

###Femap
* Open new Femap session,
* Go to File/Programming,
* Select "Open new" and choose scripts you want to open,
* Run imported script.

!!! note 
    Femap scripts can be easily embedded into the software.  
    To do so, go to Femap root folder and copy script into api folder. This will make script available under Custom Tools section.

## Contributors
<a href="https://github.com/furmanp">
  <img src="https://github.com/furmanp.png?size=50">
</a>
<a href="https://github.com/delpontis">
   <img src="https://github.com/delpontis.png?size=50">
</a>



