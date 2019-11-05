# CRDiff
CRDiff is a program that can compare Crystal reports files, and works in conjunction with your text diff tool. If your tool is configurable to use file transforms, it works even better.

# Requirements
Download and install Crystal reports for .NET 32 bit (http://downloads.businessobjects.com/akdlm/crnetruntime/clickonce/CRRuntime_32bit_13_0_22.msi)

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
1. Open CompareIt!, open Tools/Options, and select Converters. Press "Add", and specify Name: "Crystal Reports", Mask: "*.rpt", Command: "\<your path\>CRDiff.exe", Arguments: "{$Source_File} {$Converted_File}"
  
