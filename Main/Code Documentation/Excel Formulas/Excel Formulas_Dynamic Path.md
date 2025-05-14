## Dynamic Path for Project Folder
In all Excel workbooks, there is a **DynamicPath** table within the hidden **Dynamic Path** sheet that contains a path to the **Undervalued Stock Scanner** project folder.  
The Excel formula below retrieves the folder path.
```
=LEFT(CELL("filename"),SEARCH("\M",CELL("filename")))
```