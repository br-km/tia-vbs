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
'*****************************************************'
'*****************************************************'
'*****************************************************'
'Deklaracja zmiennych'
Dim Img 'As ImageFile
Dim IP 'As ImageProcess
Dim fso 'As Scripting.FileSystemObject
Dim width, height, keepAspectRatio
Dim inFilePath, inFile
Dim outFilePath, outFile
'Definicja obiektów'
Set Img = CreateObject("WIA.ImageFile")
Set IP = CreateObject("WIA.ImageProcess")
Set fso = CreateObject("Scripting.FileSystemObject")
'Ustawienie wartości zmiennych'
width = 800
height = 600
'inFilePath = "C:\Test"
inFilePath = "\\svc-625493\images\camera-photos"
inFile = fileName '"image.jpg"
outFilePath = "C:\Test-Resized"
outFile = "resized.jpg"
keepAspectRatio = false 'Zachowanie wspolczynnika proporcji'
'Zaladowanie pliku zdjecia'
Img.LoadFile inFilePath + "\" + inFile
'Definicja filtrow'
IP.Filters.Add IP.FilterInfos("Scale").FilterID
IP.Filters(1).Properties("MaximumWidth") = width
IP.Filters(1).Properties("MaximumHeight") = height
IP.Filters(1).Properties("PreserveAspectRatio") = keepAspectRatio
'Przypisanie filtrow do zdjecia'
Set Img = IP.Apply(Img)
'Sprawdzenie czy folder docelowy istnieje. Jesli nie to utworzenie'
If fso.FolderExists(outfilePath) = true Then

Else
    fso.CreateFolder outFilePath
End If
If fso.FileExists(outFilePath + "\" + outFile) Then
    fso.DeleteFile outFilePath + "\" + outFile
Else

End If
'Zapisanie zdjecia do pliku'
Img.SaveFile outFilePath + "\" + outFile
