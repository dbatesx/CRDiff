# CRDiff
CRDiff is a program that can compare Crystal reports files, and works in conjunction with your text diff tool. If your tool is configurable to use file transforms, it works even better.

# Installation
\\aus2-odx-dfs01.pdsenergy.local\devTools\CRDiff\
For faster installation, copy CRDiff folder to a local drive, then run setup.exe.
The app package is not yet signed with an authorized certificate:
You may get a message "We can't verify who created this file. Are you sure you want to run this file?" Take a chance and press "Run" ;)
You may then get a message "Publisher cannot be verified. Are you sure you want to install this application?" Press "Install"
You may then get a message "The publisher could not be verified. Are you sure you want to run this software?" Press "Run"
Your installation 

# Usage: 
  - CRDiff.exe DiffUtilPath RPTPath1 RPTPath2
    - CRDiff will serialize the two reports to json and pass the json files to your diff tool. Once you close your text diff tool, CRDiff will delete the json files.
  - CRDiff.exe RPTPath, TempPath
    - CRDiff will serialize RPTPath to TempPath (provided by your configurable text diff tool)
  - CRDiff.exe RPTPath
    - Serialize the report to a json file of the same name and location but .json extension. 
  - Parameters:
    - DiffUtilPath - Full path to external diff application .exe file that can compare two text files
    - RPTPath1 - Full path to first .rpt file to be serialized to json
    - RPTPath2 - Full path to second .rpt file to be serialized to json and compared with first file
    
# Configuration
## CompareIt!
Open CompareIt!, open Tools/Options, and select Converters. Press "Add", and specify Name: "Crystal Reports", Mask: "\*.rpt", Command: "\<CRDiff path\>CRDiff.exe", Arguments: "{$Source_File} {$Converted_File}"
## Beyond Compare
In Tools, File Formats..., press "+" and select Text Format. In the General tab, specify Mask: "\*.rpt". In the Conversion tab, specify "External program (Unicode filenames)", Loading: "\<CRDiff path>\\CRDiff %s %t", check Disable editing, then Save.
## TortoiseGit
In Settings, Diff Viewer, click Advanced, Add..., and specify Extension: ".rpt", External Program: "\<CRDiff path\>CRDiff /<path to your text compare tool\>  %base %mine"
