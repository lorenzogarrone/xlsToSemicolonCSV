Const OpenAsDefault = -2
Const FailIfNotExist = 0
Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "C:\Users\loren\Desktop\Nuova cartella (2)"

Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files
For Each objFile in colFiles
    csv_format = 6

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    src_file = objStartFolder & "\" & objFile.Name
    dest_file = objStartFolder & "\TEST\" & objFile.Name & ".csv"

    Dim oExcel
    Set oExcel = CreateObject("Excel.Application")

    Dim oBook
    Set oBook = oExcel.Workbooks.Open(src_file)

    oBook.SaveAs dest_file, csv_format

    oBook.Close False
    oExcel.Quit

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set fCSVFile = _
    oFSO.OpenTextFile(dest_file, ForReading, FailIfNotExist, OpenAsDefault)

    sFileContents = fCSVFile.ReadAll
    fCSVFile.Close
    sFileContents = Replace(sFileContents, ",",";")

    Set fCSVFile = oFSO.OpenTextFile(dest_file, ForWriting, True)
    fCSVFile.Write(sFileContents)
    fCSVFile.Close
    
    'CODED BY LORENZO GARRONE
    
Next
