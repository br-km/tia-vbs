Dim folderPath
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
folderPath = "\\svc-625493\images\camera-photos"  


Set objFolder = objFSO.GetFolder(folderPath)
Wscript.Echo objFolder.Path
Set colFiles = objFolder.Files
For Each objFile in colFiles
    Wscript.Echo objFile.Name
Next
Wscript.Echo