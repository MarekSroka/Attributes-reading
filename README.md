# Reading_Attributes project #

**1.** **VERSION 1**

 Reads attributes from a folder of *.xlsx files and generates a .xlsx report file for each folder.

**STEPS:**
1) The script looks at all Excel files in the selected folder - folder given as input parameter typed in the console.

2) Takes only rows where 'objectType' is not empty.

3) Takes all values from columns with a name containing "attributeList.attribute.name...".

4) Retrieves all values from columns where 'Attribute Value' is specified (e.g. "attributeList.attribute.string").

5) If after the column containing "attributeList.attribute.name..." there is an attributeList.attribute.<type_column> column, it retrieves the value from that column, if no value, it prints "<none>".

6) The name of the output file is the name of the folder, e.g. Component.xlsx.


**2**
Reading data from original *.xlsx files
