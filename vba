Public ppApplication As Object 'PowerPoint.Application
Public ppPresentation As Object 'PowerPoint.Presentation
Public ppObject As Object 'PowerPoint.Object
Public ppSlide As Object 'PowerPoint.Slide
Public excelOutputSheet As Worksheet
Public fileName As String
Public inputFolder As String
Public rowNumber As Long

Sub inputPath()
    inputFolderName = InputBox("フォルダのパスを入力してください")
    If inputFolderName <> 0 Then
        OutputText (inputFolderName)
    Else
        MsgBox ("入力した文字列が不正です。")
    End If
End Sub

Sub OutputText(inputFolderName)
    fileName = Dir$(inputFolderName & "\" & "*.ppt")
        
    Set ppApplication = CreateObject("PowerPoint.Application")
    Set excelOutputSheet = Workbooks.Add.Sheets(1)
    Application.ScreenUpdating = False  'VBAの速度向上のため
    
    Call inputColumnName
    rowNumber = 2   '2行目からExcelに書き込んでいくため初期値を2にする。
    
    While fileName <> ""
        Set ppPresentation = ppApplication.Presentations.Open(fileName:=inputFolderName & "\" & fileName, ReadOnly:=True)
        Call fetchObjectFromFile
        fileName = Dir$()   '変数fileNameに次のパスを入れる
    Wend
    ppPresentation.Close
    Set ppPresentation = Nothing
    ppApplication.Quit
    excelOutputSheet.Rows.AutoFit

Bye_:
    Set ppApplication = Nothing
    Set excelOutputSheet = Nothing
    Exit Sub
Err_:
    MsgBox Err.Description, vbCritical
    Resume Bye_

End Sub

Sub inputColumnName()   '一行目の項目名をExcelに書き込む
    With excelOutputSheet.Range("A1:D1")
        .Font.Bold = True
        .Value = Array("Filename", "Slide Number", "Shape Name", "Text")
    End With
End Sub

Sub fetchObjectFromFile()   'ファイルからスライドを１枚ずつ取り出し、そこから更にオブジェクトを１個ずつ取り出す
    For Each ppSlide In ppPresentation.Slides
        For Each ppObject In ppSlide.Shapes
            Call writeStringToExcel(ppObject)
        Next
    Next
End Sub

Sub writeStringToExcel(object)
    Dim columnNumber As Long
    Select Case True
    Case object.HasTextFrame
        Call writeValueToCel
        excelOutputSheet.Cells(rowNumber, "D").Value = Replace$(object.TextFrame.TextRange.Text, vbCr, vbLf)
        rowNumber = rowNumber + 1
    Case object.HasTable
        For Each Row In object.Table.Rows
            columnNumber = 4    '4列目からExcelに書き込んでいくため初期値を4にする。
            Call writeValueToCel
            For Each cell In Row.Cells
                excelOutputSheet.Cells(rowNumber, columnNumber).Value = Replace$(cell.Shape.TextFrame.TextRange.Text, vbCr, vbLf)
                columnNumber = columnNumber + 1
            Next
            rowNumber = rowNumber + 1
        Next
    Case object.HasChart
        If object.Chart.HasTitle Then
            Call writeValueToCel
            excelOutputSheet.Cells(rowNumber, "D").Value = Replace$(object.Chart.ChartTitle.Text, vbCr, vbLf)
            rowNumber = rowNumber + 1
        End If
    Case object.HasSmartArt
        For Each art In object.SmartArt.Nodes
            Call writeValueToCel
            excelOutputSheet.Cells(rowNumber, "D").Value = Replace$(art.TextFrame2.TextRange.Text, vbCr, vbLf)
            rowNumber = rowNumber + 1
        Next
    End Select
    If object.Type = msoGroup Then
        For Each object In object.GroupItems
            Call writeStringToExcel(object)
        Next
    End If
End Sub

Sub writeValueToCel()
    excelOutputSheet.Cells(rowNumber, "A").Value = fileName
    excelOutputSheet.Cells(rowNumber, "B").Value = ppSlide.SlideNumber
    excelOutputSheet.Cells(rowNumber, "C").Value = ppObject.Name
End Sub
