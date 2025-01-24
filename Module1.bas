Attribute VB_Name = "Module1"

'Subroutine to select starting directory
Sub loopAllSubFolderSlectStartDirectory()


'Intialize Variables
Dim fd As FileDialog
Dim filePath As String
Dim folderName() As String
Dim wbNew As Workbook
Dim wsNew As Worksheet
Dim newRange As Range
Dim Shop_Images As Range
Dim partNumber As String
Dim Shop As String

'Set Variables
Set wbNew = Application.Workbooks("ISAH Part Creation Assistant (version 1).xlsm")
Set wsNew = wbNew.Sheets(1)
Set newRange = wsNew.Cells(4, 2)
Set Shop_Images = wsNew.Cells(10, 1)
Set fd = Application.FileDialog(msoFileDialogFolderPicker)

'Call Initialize Flags function to reset all flags
Call InitializeFlags(3, 9)

'Collect the part number
If fd.Show = True Then
    filePath = fd.SelectedItems(1)
    folderName = Split(filePath, "\")
    partNumber = folderName(UBound(folderName))
    
    With newRange
        .Value = partNumber
         .Font.Name = "Gotham Black"
         .Font.Bold = True
         .Font.Size = 26
    End With
    
    'Pass the part number to the subroutine
    Call LoopAllSubFolders(fd.SelectedItems(1), partNumber)
Else
    'No file was picked, cancel and clear
    wsNew.Cells(11, 1).Value = ""
    wsNew.Cells(11, 7).Value = ""
    wsNew.Cells(20, 1).Value = ""
    wsNew.Cells(20, 7).Value = ""

End If

'Compare revisions
Call RevComparison


'Clears the sections if they are empty
ClearShopImagesCells (partNumber)
ClearProductionFileCells (partNumber)
ClearProductImageCells (partNumber)
ClearDrawingCells (partNumber)

End Sub
'Sub that loops through a folders sub folders
Sub LoopAllSubFolders(ByVal folderPath As String, partNum As String)

'Initialize Variable Names
Dim fileName As String
Dim filePath As String
Dim numFolders As Long
Dim folders() As String
Dim i As Long
Dim fileSystem As Object
Set fileSystem = CreateObject("Scripting.FileSystemObject")
Dim file As Object
Dim Range As Range
Dim strText As String
Dim wsNew As Worksheet
Dim folderSystem As Object
Dim folderName As Object
Dim wbNew As Workbook
Dim newRange As Range
Dim RevNumber As String
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim extentions() As String
Dim files() As String
Dim v As Integer
Dim partNumber As Variant
Dim ValChanged As Integer
Dim prodFilesRange As Range
Dim ProdImages As Range
Dim Drawings As Range
Dim Name As String

'Pass values into variables
c = 0
b = 0
a = 0
d = 0
ValChanged = 0
Set wbNew = Application.Workbooks("ISAH Part Creation Assistant (version 1)")
Set wsNew = wbNew.Sheets(1)
Set Range = wsNew.Cells(3, 2)
Set newRange = wsNew.Cells(11, 1)
Set prodFilesRange = wsNew.Cells(20, 1)
Set ProdImages = wsNew.Cells(11, 7)
Set Drawings = wsNew.Cells(20, 7)



'Parse the file name
If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
fileName = Dir(folderPath & "*.*", vbDirectory)

'While the file exists
While Len(fileName) <> 0

    If Left(fileName, 1) <> "." Then
    
    'Get the full file path
    fullFilePath = folderPath & fileName
    
    
    If (GetAttr(fullFilePath) And vbDirectory) = vbDirectory Then
        ReDim Preserve folders(0 To numFolders) As String
        folders(numFolders) = fullFilePath
        numFolders = numFolders + 1
    Else
        
        Set file = fileSystem.GetFile(fullFilePath)
        
        'Check if the parent folder is Shop Images
        If (UCase(file.ParentFolder.Name) = "SHOP_IMAGES") Or (UCase(file.ParentFolder.Name) = "SHOP IMAGES") Or (UCase(file.ParentFolder.Name) = "SHOP IMAGES") Or (UCase(file.ParentFolder.Name) = "SHOP IMAGE") Or (UCase(file.ParentFolder.Name) = "SHOP_IMAGE") Then
                'Throw error if there are too many
                If a > 6 Then
                    MsgBox "There are too many files in the Shop_Images folder. That isn't possible. Fix it and try again", vbCritical, "Too Many Files"
                    Exit Sub
                End If
                
                'Split the file path ad parse different items
                extentions = Split(folderPath & fileName, ".")
                
                'Original File Path
                newRange.Offset(0 + a, 0).Value = folderPath & fileName
                
                'New File path with different extention
                newRange.Offset(0 + a, 1).Value = "\\lr-app01\d\" & Right(folderPath, Len(folderPath) - 3) & fileName
                
                'File name
                newRange.Offset(0 + a, 2).Value = Left(fileName, -1 + Len(fileName) - Len(extentions(UBound(extentions))))
                
                'File Extension
                newRange.Offset(0 + a, 3).Value = extentions(UBound(extentions))
            
                'Call return RevNumber function to parse rev from file name
                RevNumber = ReturnRevNumber(Left(fileName, -1 + Len(fileName) - Len(extentions(UBound(extentions)))))
            
            'If Rev hasn't been updated, update it
            If (Not (RevNumber = "")) And ValChanged = 0 Then
                With wsNew.Cells(9, 5)
                    .Value = UCase(RevNumber)
                    .Font.Size = 26
                    .Font.Bold = True
                    ValChanged = 1
                End With
                ElseIf Not (ValChanged > 0) Then
                    With wsNew.Cells(9, 5)
                    .Value = "N/A"
                    .Font.Size = 26
                    .Font.Bold = True
                    End With
            End If
            
            
            'Increment file count
            a = a + 1
            wsNew.Cells(9, 3).Value = a
            
            'Check for Production Files Sub Folder
        ElseIf (UCase(file.ParentFolder.Name) = "PRODUCTION FILES") Or (UCase(file.ParentFolder.Name) = "PRODUCTION_FILES") Or (UCase(file.ParentFolder.Name) = "PRODUCTION_FILE") Or (UCase(file.ParentFolder.Name) = "PRODUCTION FILE") Then
                    'If theres more than 11 files, abort operation
                    If b > 11 Then
                        MsgBox "There are too many files in the Production Files folder. May need a manual check", vbCritical, "Too Many Files"
                        Exit Sub
                    End If
                    
                    'Split File psth into separate items
                    extentions = Split(folderPath & fileName, ".")
                    'Original path
                    prodFilesRange.Offset(0 + b, 0).Value = folderPath & fileName
                    'Updated path
                    prodFilesRange.Offset(0 + b, 1).Value = "\\lr-app01\d\" & Right(folderPath, Len(folderPath) - 3) & fileName
                    'File Name
                    prodFilesRange.Offset(0 + b, 2).Value = Left(fileName, -1 + Len(fileName) - Len(extentions(UBound(extentions))))
                    'File extension
                    prodFilesRange.Offset(0 + b, 3).Value = extentions(UBound(extentions))
                 
                 'count number of files in folder
                 b = b + 1
                 wsNew.Cells(18, 3).Value = b
                 
                    
        'Check for Product Images Sub Folder
        ElseIf (UCase(file.ParentFolder.Name) = "PRODUCT IMAGES") Or (UCase(file.ParentFolder.Name) = "PRODUCT_IMAGES") Or (UCase(file.ParentFolder.Name) = "PRODUCT_IMAGE") Or (UCase(file.ParentFolder.Name) = "PRODUCT IMAGE") Then
                    'If theres more than 3 files, abort operation
                    If c > 3 Then
                        MsgBox "There are too many files in the Product Images folder. May need a manual check", vbCritical, "Too Many Files"
                        Exit Sub
                    End If
                    
                    'Split file into seperate items
                    extentions = Split(folderPath & fileName, ".")
                    'Original Path
                    ProdImages.Offset(0 + c, 0).Value = folderPath & fileName
                    'Updated Path
                    ProdImages.Offset(0 + c, 1).Value = "\\lr-app01\d\" & Right(folderPath, Len(folderPath) - 3) & fileName
                    'File Name
                    ProdImages.Offset(0 + c, 8).Value = Left(fileName, -1 + Len(fileName) - Len(extentions(UBound(extentions))))
                    'File Extension
                    ProdImages.Offset(0 + c, 10).Value = extentions(UBound(extentions))
                    
                    'Parse revision number
                    RevNumber = ReturnRevNumberImage(Left(fileName, -1 + Len(fileName) - Len(extentions(UBound(extentions)))))
                
                
                'If Rev not changed, update it
                If (Not (RevNumber = "")) And ValChanged = 0 Then
            
                With wsNew.Cells(9, 20)
                    .Value = UCase(RevNumber)
                    .Font.Size = 26
                    .Font.Bold = True
                    ValChanged = 1
                End With
                ElseIf Not (ValChanged > 0) Then
                    With wsNew.Cells(9, 20)
                    .Value = "N/A"
                    .Font.Size = 26
                    .Font.Bold = True
                    End With
            End If
                
                'Count number of files
                 c = c + 1
                 wsNew.Cells(9, 17).Value = c
                 
            'Check for drawings folder
            ElseIf (UCase(file.ParentFolder.Name) = "CURRENT_DRAWINGS") Or (UCase(file.ParentFolder.Name) = "CURRENT DRAWINGS") Or (UCase(file.ParentFolder.Name) = "CURRENT DRAWING") Or (UCase(file.ParentFolder.Name) = "CURRENT_DRAWING") Then
                    'If more than 3 files, abort
                    If d > 3 Then
                        MsgBox "There are too many files in the Drawings folder. May need a manual check", vbCritical, "Too Many Files"
                        Exit Sub
                    End If
                    
                    'Split file path into seperate items
                    extentions = Split(folderPath & fileName, ".")
                    'Original Path
                    Drawings.Offset(0 + d, 0).Value = folderPath & fileName
                    'Updated Path
                    Drawings.Offset(0 + d, 1).Value = "\\lr-app01\d\" & Right(folderPath, Len(folderPath) - 3) & fileName
                    'File name
                    Drawings.Offset(0 + d, 8).Value = Left(fileName, -1 + Len(fileName) - Len(extentions(UBound(extentions))))
                    'File Extention
                    Drawings.Offset(0 + d, 10).Value = extentions(UBound(extentions))
                   
                   'Parse revision number
                    RevNumber = ReturnRevNumberDrawing(Left(fileName, -1 + Len(fileName) - Len(extentions(UBound(extentions)))))
                    
                 'If revision number unchanged, change it
                If (Not (RevNumber = "")) And ValChanged = 0 Then
                
                
                With wsNew.Cells(18, 20)
                    .Value = UCase(RevNumber)
                    .Font.Size = 26
                    .Font.Bold = True
                    ValChanged = 1
                End With
                ElseIf Not (ValChanged > 0) Then
                    With wsNew.Cells(18, 20)
                    .Value = "N/A"
                    .Font.Size = 26
                    .Font.Bold = True
                    End With
            End If
                
                'Cunt number of files
                 d = d + 1
                 wsNew.Cells(18, 17).Value = d
                      
        End If
    End If
End If
    
    fileName = Dir()

    
    
Wend
'Loop through folders sub folders and set off flags if folder exists
For i = 0 To numFolders - 1

files = Split(folders(i), "\")
Name = UCase(files(UBound(files)))
    If (Name = "SHOP_IMAGES") Or (Name = "SHOP IMAGES") Or (Name = "SHOP IMAGES") Or (Name = "SHOP IMAGE") Or (Name = "SHOP_IMAGE") Then
    'Shop Images Folder exists
            With wsNew.Cells(5, 9)
                        .Interior.Color = vbGreen
                        .Value = "EXISTS"
                    End With
    ElseIf (Name = "PRODUCTION FILES") Or (Name = "PRODUCTION_FILES") Or (Name = "PRODUCTION_FILE") Or (Name = "PRODUCTION FILE") Then
    'Production Images Folder exists
                 With wsNew.Cells(4, 9)
                        .Interior.Color = vbGreen
                        .Value = "EXISTS"
                    End With
    ElseIf (Name = "CURRENT_DRAWINGS") Or (Name = "CURRENT DRAWINGS") Or (Name = "CURRENT DRAWING") Or (Name = "CURRENT_DRAWING") Then
    'Current Drawings Folder exists
                With wsNew.Cells(6, 9)
                        .Interior.Color = vbGreen
                        .Value = "EXISTS"
                    End With
    ElseIf (Name = "PRODUCT IMAGES") Or (Name = "PRODUCT_IMAGES") Or (Name = "PRODUCT_IMAGE") Or (Name = "PRODUCT IMAGE") Then
      'Product Images Folder Exists
        With wsNew.Cells(3, 9)
                        .Interior.Color = vbGreen
                        .Value = "EXISTS"
                    End With
    End If
    
    'Loop into the subfolder
    LoopAllSubFolders folders(i), partNum
    
Next i

End Sub
'Function to parse rev number from shop image file name
Function ReturnRevNumber(fileName As String) As String

'Initialize variables
Dim strText As String
Dim Rev As String

'Set Variables
strText = LCase(fileName)

'Parse name by finding where the text "rev" is, where the text "shop" is, and systematically getting the info between. The revision number is always between "Rev" and "-shop"
If (InStr(1, strText, "rev") > 0) Then
    
    Rev = Right(strText, 1 + Len(strText) - InStr(1, strText, "rev", vbTextCompare))
    Rev = Right(Rev, Len(Rev) - 3)
    If (InStr(1, Rev, "shop", vbTextCompare) > 0) Then

        Rev = Left(Rev, InStr(1, Rev, "-shop", vbTextCompare) - 1)
        ReturnRevNumber = Rev
    End If

End If
End Function
'Function to parse rev number from Product Image file name
Function ReturnRevNumberImage(fileName As String) As String


'Intialize Variables
Dim strText As String
Dim Rev As String

'Set variables
strText = LCase(fileName)

'Parse name by finding where the text "rev" is, where the text "prod" is, and systematically getting the info in between
If (InStr(1, strText, "rev") > 0) Then
    Rev = Right(strText, 1 + Len(strText) - InStr(1, strText, "rev", vbTextCompare))
    Rev = Right(Rev, Len(Rev) - 3)
    If (InStr(1, Rev, "prod", vbTextCompare) > 0) Then

        Rev = Left(Rev, InStr(1, Rev, "-prod", vbTextCompare) - 1)
        ReturnRevNumberImage = Rev
    End If

End If
End Function
'Funciton to parse rev number from Drawing file name
Function ReturnRevNumberDrawing(fileName As String) As String

'Initialize Variables
Dim strText As String
Dim Rev As String

'Set Variables
strText = LCase(fileName)

'Parse out the revision number by finding where "rev" is, the revision number is usually just next to "Rev" with nothing next to it
If (InStr(1, strText, "rev") > 0) Then

    'If the word "Drawing" is present, get the number in between "rev" and "-drawing", otherwise just get whatever's nex to "rev"
    If (InStr(1, strText, "drawing") > 0) Or (InStr(1, strText, "-drawing") > 0) Then
        Rev = Left(strText, Len(strText) - Len(" DRAWING"))

        Rev = Right(Rev, Len(Rev) - InStr(1, Rev, "rev", vbTextCompare) + 1)

        Rev = Right(Rev, Len(Rev) - 3)

        ReturnRevNumberDrawing = Rev
    Else
        Rev = Right(strText, Len(strText) - InStr(1, strText, "rev", vbTextCompare) + 1)

        Rev = Right(Rev, Len(Rev) - 3)

        ReturnRevNumberDrawing = Rev
    End If
    
        
Else
    Rev = "N/A"

End If
End Function
'Function to clear out the cells if there is nothing there and update the flags
Function ClearShopImagesCells(partNumber As String)

'Intialize variables
Dim wbNew As Workbook
Dim wsNew As Worksheet
Dim Shop_Images As Range

'Set the variables
Set wbNew = Application.Workbooks("ISAH Part Creation Assistant (version 1)")
Set wsNew = wbNew.Sheets(1)
Set Shop_Images = wsNew.Cells(11, 1)

'If the part number is not in the file path, remove it
For v = 0 To 6
    Shop = Shop_Images.Offset(0 + v, 0).Value
    If (InStr(Shop, partNumber)) = 0 Then
        Shop_Images.Offset(0 + v, 2).Value = ""
        Shop_Images.Offset(0 + v, 1).Value = ""
        Shop_Images.Offset(0 + v, 0).Value = ""
        Shop_Images.Offset(0 + v, 3).Value = ""
    End If
'If the first cell is empty, clear out the whole section and update the flag to be empty
    If wsNew.Cells(11, 1).Value = "" Then
        With wsNew.Cells(5, 10)
            .Interior.Color = vbRed
            .Value = "EMPTY"
        End With
        wsNew.Cells(9, 3).Value = "N/A"
        wsNew.Cells(9, 5).Value = "N/A"
        For i = 0 To 6
        Shop_Images.Offset(0 + i, 2).Value = ""
        Shop_Images.Offset(0 + i, 1).Value = ""
        Shop_Images.Offset(0 + i, 0).Value = ""
        Shop_Images.Offset(0 + i, 3).Value = ""
        Next i
        Exit Function
    Else
        With wsNew.Cells(5, 10)
            .Interior.Color = vbGreen
            .Value = "NOT EMPTY"
        End With
    End If
    
Next v

End Function
'Function to clear out the cells if there is nothing and update the flags
Function ClearProductionFileCells(partNumber As String)

'Initialie Variables
Dim wbNew As Workbook
Dim wsNew As Worksheet
Dim ProdFiles As Range

'Set Variables
Set wbNew = Application.Workbooks("ISAH Part Creation Assistant (version 1)")
Set wsNew = wbNew.Sheets(1)
Set ProdFiles = wsNew.Cells(20, 1)


'Clear row if the part number is not in the path
For v = 0 To 11
    Prod = ProdFiles.Offset(0 + v, 0).Value
    If (InStr(Prod, partNumber)) = 0 Then
        ProdFiles.Offset(0 + v, 2).Value = ""
        ProdFiles.Offset(0 + v, 1).Value = ""
        ProdFiles.Offset(0 + v, 0).Value = ""
        ProdFiles.Offset(0 + v, 3).Value = ""
    End If
    
'If first cell is empty, clear section and update flag
    If wsNew.Cells(20, 1).Value = "" Then
        With wsNew.Cells(4, 10)
            .Interior.Color = vbRed
            .Value = "EMPTY"
        End With
        wsNew.Cells(18, 3).Value = "N/A"
        
        For i = 0 To 11
        ProdFiles.Offset(0 + i, 2).Value = ""
        ProdFiles.Offset(0 + i, 1).Value = ""
        ProdFiles.Offset(0 + i, 0).Value = ""
        ProdFiles.Offset(0 + i, 3).Value = ""
        Next i
        Exit Function
    Else
        With wsNew.Cells(4, 10)
            .Interior.Color = vbGreen
            .Value = "NOT EMPTY"
            
                  
        End With
    End If
    
Next v


End Function

'Function to clear out the cells if there is nothing and update the flags
Function ClearProductImageCells(partNumber As String)

'Initialize Variables
Dim wbNew As Workbook
Dim wsNew As Worksheet
Dim ProdImages As Range

'Set Variables
Set wbNew = Application.Workbooks("ISAH Part Creation Assistant (version 1)")
Set wsNew = wbNew.Sheets(1)
Set ProdImages = wsNew.Cells(11, 7)




'If part number is not in the file path, clear out the row
For v = 0 To 2
    Prod = ProdImages.Offset(0 + v, 0).Value
    If (InStr(Prod, partNumber)) = 0 Then
        ProdImages.Offset(0 + v, 8).Value = ""
        ProdImages.Offset(0 + v, 1).Value = ""
        ProdImages.Offset(0 + v, 0).Value = ""
        ProdImages.Offset(0 + v, 10).Value = ""
    End If
    
    'If the first cell is empty, clear out the section and update the flags
    If wsNew.Cells(11, 7).Value = "" Then
        With wsNew.Cells(3, 10)
            .Interior.Color = vbRed
            .Value = "EMPTY"
        End With
        wsNew.Cells(9, 17).Value = "N/A"
        wsNew.Cells(9, 20).Value = "N/A"
        For i = 0 To 2
        ProdImages.Offset(0 + i, 8).Value = ""
        ProdImages.Offset(0 + i, 1).Value = ""
        ProdImages.Offset(0 + i, 0).Value = ""
        ProdImages.Offset(0 + i, 10).Value = ""
        Next i
        Exit Function
    Else
        With wsNew.Cells(3, 10)
            .Interior.Color = vbGreen
            .Value = "NOT EMPTY"
        End With
    End If
    
Next v


End Function
'Function to clear out the cells if there is nothing and update the flags
Function ClearDrawingCells(partNumber As String)
'Initialize the variables
Dim wbNew As Workbook
Dim wsNew As Worksheet
Dim Drawings As Range

'Set the variables
Set wbNew = Application.Workbooks("ISAH Part Creation Assistant (version 1)")
Set wsNew = wbNew.Sheets(1)
Set Drawings = wsNew.Cells(20, 7)




'If Part number is not in the file path, clear the row
For v = 0 To 2
    Prod = Drawings.Offset(0 + v, 0).Value
    If (InStr(Prod, partNumber)) = 0 Then
        Drawings.Offset(0 + v, 8).Value = ""
        Drawings.Offset(0 + v, 1).Value = ""
        Drawings.Offset(0 + v, 0).Value = ""
        Drawings.Offset(0 + v, 10).Value = ""
    End If
    'If first cell is empty, clear the whole section and update the flags
    If wsNew.Cells(20, 7).Value = "" Then
        With wsNew.Cells(6, 10)
            .Interior.Color = vbRed
            .Value = "EMPTY"
        End With
        wsNew.Cells(18, 17).Value = "N/A"
        wsNew.Cells(18, 20).Value = "N/A"
        For i = 0 To 2
            Drawings.Offset(0 + i, 8).Value = ""
            Drawings.Offset(0 + i, 1).Value = ""
            Drawings.Offset(0 + i, 0).Value = ""
            Drawings.Offset(0 + i, 10).Value = ""
        Next i
        Exit Function
    Else
        With wsNew.Cells(6, 10)
            .Interior.Color = vbGreen
            .Value = "NOT EMPTY"
            
            
            
        End With
    End If
    
Next v


End Function
'Initialize the flags to be empty and not exist, until proven otherwise
Function InitializeFlags(RowNum As Integer, ColNum As Integer)

'initialize variables
Dim wbNew As Workbook
Dim wsNew As Worksheet
Dim Flags As Range

'Set the flags default
Set wbNew = Application.Workbooks("ISAH Part Creation Assistant (version 1).xlsm")
Set wsNew = wbNew.Sheets(1)
Set Flags = wsNew.Cells(RowNum, ColNum)

For i = 0 To 3
    With Flags.Offset(i, 0)
        .Value = "DOESN'T EXIST"
        .Interior.Color = vbRed
        
    End With
Next i
For v = 0 To 3
    With Flags.Offset(v, 1)
        .Value = "EMPTY"
        .Interior.Color = vbRed
        
    End With
Next v
End Function

'Subroutine to compare rev numbers
Sub RevComparison()


'Initialize Variables
Dim wbNew As Workbook
Dim wsNew As Worksheet
Dim ProdImageRev As Variant
Dim DrawingRev As Variant
Dim ShopImageRev As Variant
Dim ProdImageRevMain As Variant
Dim DrawingRevMain As Variant
Dim ShopImageRevMain As Variant
Dim DrawingMain() As String
Dim ProdMain() As String
Dim ShopMain() As String

'Set variables
Set wbNew = Application.Workbooks("ISAH Part Creation Assistant (version 1).xlsm")
Set wsNew = wbNew.Sheets(1)

Dim RevComp As Range
Set RevComp = wsNew.Cells(4, 20)

'Check if Rev is N/A or an actual value
If Not (wsNew.Cells(18, 20).Value = "N/A") Then
    DrawingRev = wsNew.Cells(18, 20).Value
    RevComp.Offset(0, 0).Value = DrawingRev
Else
    RevComp.Offset(0, 0).Value = "N/A"
End If

If Not (wsNew.Cells(9, 20).Value = "N/A") Then
    ProdImageRev = wsNew.Cells(9, 20).Value
    RevComp.Offset(1, 0).Value = ProdImageRev
Else
    RevComp.Offset(1, 0).Value = "N/A"
End If

If Not (wsNew.Cells(9, 5).Value = "N/A") Then
    ShopImageRev = wsNew.Cells(9, 5).Value
    RevComp.Offset(2, 0).Value = ShopImageRev

Else
    RevComp.Offset(2, 0).Value = "N/A"
End If



'Parse and compare the first part of the rev number, if the first part is the same, revs are considered similar, if the first and second part are the same, they are consistent, if even a single one is different, it's inconsistent,
'If they are all N/A, Rev is considered N/A
If (InStr(1, DrawingRev, ".", vbTextCompare) > 0) Then
   DrawingMain = Split(DrawingRev, ".")
    DrawingRevMain = DrawingMain(LBound(DrawingMain()))
Else
    DrawingRevMain = DrawingRev
End If
If (InStr(1, ProdImageRev, ".", vbTextCompare) > 0) Then
   ProdMain = Split(ProdImageRev, ".")
    ProdImageRevMain = ProdMain(LBound(ProdMain()))
Else
    ProdImageRevMain = ProdImageRev

End If
If (InStr(1, ShopImageRev, ".", vbTextCompare) > 0) Then
   ShopMain = Split(ShopImageRev, ".")
    ShopImageRevMain = ShopMain(LBound(ShopMain()))
Else
    ShopImageRevMain = ShopImageRev
End If

'Compare
If CStr(DrawingRev) = CStr(ProdImageRev) And CStr(DrawingRev) = CStr(ShopImageRev) Then
  If CStr(DrawingRev) = "" Then
    With RevComp.Offset(0, 1)
            .Value = "There are no recorded rev numbers"
            .Interior.ColorIndex = 36
        End With
  Else
        With RevComp.Offset(0, 1)
            .Value = "The Rev numbers for all documents are the same, this part is consistent"
            .Interior.Color = vbGreen
        End With
    End If
Else
    If CStr(DrawingRevMain) = CStr(ProdImageRevMain) And CStr(DrawingRevMain) = CStr(ShopImageRevMain) Then
      

            With RevComp.Offset(0, 1)
                .Value = "The Rev numbers for all documents are very similar, checkover to ensure consistency"
                .Interior.ColorIndex = 35
            End With

    Else
        With RevComp.Offset(0, 1)
            .Value = "The Rev numbers for all documents are not the same, this part is not consistent and files should be looked over"
            .Interior.Color = vbRed
        End With
        End If
End If

End Sub
