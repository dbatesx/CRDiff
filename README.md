# CRDiff
CRDiff is a program that can serialize the binary Crystal reports .rpt files to human readable .json files, and works in conjunction with your text diff tool. If your tool is configurable to use file transforms, it works even better.

# Installation
The published location is \\\\aus2-odx-dfs01.pdsenergy.local\\devTools\\CRDiff

Copy the contents to the local folder of your choice (\<CRDiff Path\>) - Having that folder in your PATH environment variable can be helpful.

You may get a message "Are you sure you want to run this software?" 
* Uncheck "Always ask before opening this file" and 
* press "Run"

If when CRDiff.exe is run with no parameters, you don't see a valid CR Version displayed, you may need to install the Crystal Reports Runtime, available in the Crystal Reports Runtime folder.

# Usage: 
  - CRDiff.exe DiffAppPath ReportFilename1 ReportFilename2
    - CRDiff will serialize the two reports to json and pass the json files to your diff tool. Once you close your text diff tool, CRDiff will delete the json files (if they hadn't already existed).
  - CRDiff.exe ReportFilename, TempFilename
    - CRDiff will serialize RPTPath to TempPath, provided by your configurable text diff tool, which is responsible for cleaning up after itself.
  - CRDiff.exe ReportFilename
    - Serialize the report to a json file of the same name and location but with .json extension. 
  - Parameters:
    - DiffAppPath - path to external differencing application .exe file that can compare two text files
    - ReportFilename1 - path to first .rpt file to be serialized to json, or it can be an existing report.json file
    - ReportFilename2 - path to second .rpt file to be serialized to json (or existing .json file) and compared with first file
    - TempFilename - path to a temporary file that will likely be deleted after use.
    
# Configuration
## CompareIt!
Open CompareIt!, open Tools/Options, and select Converters. Press "Add", and specify Name: "Crystal Reports", Mask: "\*.rpt", Command: "\<CRDiff path\>CRDiff.exe", Arguments: "{$Source_File} {$Converted_File}"
## Beyond Compare
In Tools, File Formats..., press "+" and select Text Format. In the General tab, specify Mask: "\*.rpt". In the Conversion tab, specify "External program (Unicode filenames)", Loading: "\<CRDiff path\>CRDiff.exe %s %t", check Disable editing, then Save.
## TortoiseGit
In Settings, Diff Viewer, click Advanced, Add..., and specify Extension: ".rpt", External Program: "\<CRDiff path\>CRDiff.exe \<path to your text compare tool\>  %base %mine"
