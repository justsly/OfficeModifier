# OfficeModifier
OfficeModifier is technically based on the great tool OfficePurge from Andrew Oliveau (@AndrewOliveau).


I build OfficeModifier to easily add and delete Streams into CFBF files.

The tool can be like this:

# ARGUMENTS/OPTIONS
* <b>-f </b> - Document filename to modify 
* <b>-n </b> - Stream name inside the CFBF File
* <b>-a </b> - Path to Stream to add into the CFBF File 
* <b>-r </b> - Stream which should be deleted (ex. Table1) 
* <b>-h </b> - Show help menu

# EXAMPLES

`OfficeModifier.exe -f example.docm -n StreamNameInCFBFStructure -a C:\PathTo\StreamOnDisk`
`OfficeModifier.exe -f example.docm -r Table1`
