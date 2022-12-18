# msidump
commandline utility for dumping MSI database information of the specified MSI file.

### Usage:
` msidump.exe [/f] msiPath [/t table1;table2...] [/l table;store] [/x xslFilePath] [/o outputDirectory] `

**Options:**

- ***msiPath*** - MSI file to dump. (Required)
- ***[/l table|store]***       list of table or storage names. (Optional) 
- ***[/t table1;table2...]***     MSI tables to dump. (Optional) 
- ***[/a store1;store2...]***     apply specified embedded MSTs. (Optional) 
- ***[/e mstfile1;mstfile2...]***  apply specified external MST file/s. (Optional) 
- ***[/x xslFilePath]***         XML Stylesheet file path. (Optional) 
- ***[/n outputFileName]***    output filename. (Optional) 
- ***[/o outputDirectory]***     output directory. (Optional) 
- ***[/b]***                     Suppress summary information stream dump. 
- ***[/d n1=path1 n2=path2...]*** diffMode: compares two msidump XML files. n1, n2 should be unique short names. 


**Example:**

`$>msidump c:\tmp\mps.msi `
>      dumps all tables and transform views to the same directory as the msi file.
`$>msidump /d RTM=mps_rtm.msi.mxl B05=mps_b05.msi.mxl `
>      dumps all the changes made on the tables between RTM and B05.


