# Reading_Attributes project #

**A** **VERSION 1**

 Reads attributes from a folder of *.xlsx preprocessed files and generates a .xlsx report file for each folder.

**STEPS:**
  1) The script looks at all Excel files in the selected folder - folder given as input parameter typed in the console.

  2) Takes only rows where 'objectType' is not empty.

  3) Takes all values from columns with a name containing "attributeList.attribute.name...".

  4) Retrieves all values from columns where 'Attribute Value' is specified (e.g. "attributeList.attribute.string").

  5) If after the column containing "attributeList.attribute.name..." there is an attributeList.attribute.<type_column> column, it retrieves the value from that column, if no value, it prints "<none>".

  6) The name of the output file is the name of the folder, e.g. Component.xlsx.


**B**
Reads data from original *.xml files

STEPS:
   1) The script looks at all *.xmls files in the selected folder - folder given as input parameter typed in the console
      
   2) Takes only rows where 'objectType' is not empty.

   3) Creates sheets that cluster files with the same object_type (sheets name = object_type)
   
   4) Each sheets has a column with the file name 
   
   5) Creates a matrix with all types of attributes available for a given object_type in the column names
   
   6) Attribute names are the values of the matrix
   
   7) Prints a summary with the number of files and file names that have been pulled into the report
   
   8) Generates XML_Attributes.xlsx file
   
   9) In the Report.xlsx A pivot table is then automatically created based on the sheets in the XML_Attributes.xlsx file, where we can check the attributes report 
   






