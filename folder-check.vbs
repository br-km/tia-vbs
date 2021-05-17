Dim folderPath,numFiles, outPathForFile
Dim objFSO, objFolder, objFiles
Dim latestDate, tempDate
Dim deletedFiles
Dim fileName
'ścieżka docelowa dla pliku'
outPathForFile = "C:\"
'ścieżka do folderu ze zdjęciami
folderPath = "\\svc-625493\images\camera-photos"
 
'zdefiniowanie obiektów'
Set objFSO = CreateObject("Scripting.FileSystemObject")  
Set objFolder = objFSO.GetFolder(folderPath)  

'przypisanie wartości ilości plików w folderze'
numFiles = objFolder.Files.Count

'wyszukiwanie najnowszego pliku i kasowanie wszystkich starych'
Do While numFiles > 1
    latestDate = 0
    Set objFiles = objFolder.Files
    For Each file in objFiles
        tempDate = file.DateCreated
        If CDate(tempDate)<CDate(latestDate) Or latestDate = 0 Then
            latestDate = tempDate
        End If
    Next
    For Each file in objFiles
        If file.DateCreated = latestDate Then
            file.Delete true
            deletedFiles = deletedFiles + 1
            numFiles = numFiles - 1
        End if 
    Next
Loop
Wscript.Echo deletedFiles
'Nazwa pliku który został'
Wscript.Echo objFolder.Path
Set FolderFiles = objFolder.Files
For Each File in FolderFiles
    fileName = File.Name
    Wscript.Echo fileName
Next