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
width = 100
height = 100
'inFilePath = "C:\Test"
inFilePath = "\\svc-625493\images\camera-photos"
inFile = "image.jpg"
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
