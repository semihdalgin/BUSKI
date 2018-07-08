Sub BUSKI()
'
' Semih Dalgin
'
' Klavye Kısayolu: Ctrl+w
'
Dim sd As Integer
Dim sd2 As Integer
Dim say As Integer
Dim say3 As Integer
Dim Ay(1 To 13) As String
Dim FirstCellWithText As Range
Dim Renk(1 To 2) As Long
Dim ABC As Integer
Dim fark As Integer

Dim mainworkBook As Workbook
Set mainworkBook = ActiveWorkbook
Dim i As Integer
Dim AD As String
Dim sem As String
Dim se As Integer
Dim currRPT As String
Dim tarih As String
Dim asd
Dim asdasd

asd = Now
asdasd = Year(asd)


tarih = Str(asdasd)

Getthename = mainworkBook.ActiveSheet.Name

ABC = 1

Renk(1) = 15773696
Renk(2) = 16776960

say = 4
sd = 0
sd2 = 1
sd3 = 0

Dim AA As Boolean

AA = True
Ay(1) = "OCAK"
Ay(2) = "ŞUBAT"
Ay(3) = "MART"
Ay(4) = "NİSAN"
Ay(5) = "MAYIS"
Ay(6) = "HAZİRAN"
Ay(7) = "TEMMUZ"
Ay(8) = "AĞUSTOS"
Ay(9) = "EYLÜL"
Ay(10) = "EKİM"
Ay(11) = "KASIM"
Ay(12) = "ARALIK"
Ay(13) = ""

Sheets.Add

Dim adim
adim = Sheets.Count

Sheets.Add After:=Sheets(adim)
i = mainworkBook.Sheets.Count
ActiveSheet.Select
currRPT = Left((ActiveSheet.Name), 5)
sem = "Sayfa"

If currRPT = sem Then
    AD = "Sayfa" & i - 1
    If WorksheetExists2("BUSKI") Then
        If WorksheetExists2(AD) Then
        Sheets(AD).Select
        ActiveWindow.SelectedSheets.Delete
        End If
    Else
    Sheets(AD).Name = "BUSKI"
    End If
Else
    AD = "Sheet" & i - 1
    If WorksheetExists2("Sheet1") Then
    Sheets(AD).Select
    ActiveWindow.SelectedSheets.Delete
    Else
    Sheets(AD).Name = "BUSKI"
    End If

End If
Sheets("BUSKI").Select

' Format !!**********************************************************************************************************************
'
Range("A1:Q1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = tarih + " YILI BUSKİ ADINA İRTİFAK HAKKI KURULAN TAŞINMAZLAR"
    Range("A2:O2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveWindow.SmallScroll ToRight:=8
    Range("P2:Q2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("P3:P4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("Q3:Q4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("I3:O3").Select
    Range("O3").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A3:H3").Select
    Range("H3").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "MÜLKİYET BİLGİLERİ"
    Range("I3:O3").Select
    ActiveCell.FormulaR1C1 = "İRTİFAK"
    Range("P3:P4").Select
    ActiveCell.FormulaR1C1 = "KAYIDI" & Chr(10) & "TARİHİ"
    Range("Q3:Q4").Select
    ActiveCell.FormulaR1C1 = "AÇIKLAMA"
    Range("Q5").Select
    Range("P2:Q2").Select
    ActiveCell.FormulaR1C1 = "BUSKİ MÜLKİYETİNDEN ÇIKIŞI"
    Range("A4").Select
    ActiveWindow.Zoom = 50
    ActiveWindow.Zoom = 60
    ActiveWindow.SmallScroll Down:=-18
    ActiveWindow.SmallScroll ToRight:=1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "AY"
    ActiveWindow.LargeScroll Down:=1
    Range("A27").Select
    ActiveWindow.SmallScroll Down:=-48
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "E.B.S." & Chr(10) & "I.D."
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "İLÇE"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "MAHALLE/KÖY"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "ADA"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "PARSEL"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "ALAN"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "CİNSİ"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "KİMDEN ALINDI"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "İRTİFAK ALAN" & Chr(10) & "(m2)"
    With ActiveCell.Characters(Start:=1, Length:=15).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=16, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = True
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=17, Length:=1).Font
        .Name = "Calibri"
        .FontStyle = "Normal"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "HİSSE"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "İRTİFAK DEĞERİ" & Chr(10) & "(TL)"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "EDİNME TARİHİ"
    Range("N4").Select
    ActiveWindow.SmallScroll ToRight:=8
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "KULLANIM AMACI"
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "BÜTÇE KODU"
    Range("A1:Q4").Select
    Range("Q3").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Columns("Q:Q").ColumnWidth = 15.71
    Columns("Q:Q").ColumnWidth = 17.14
    Columns("O:O").ColumnWidth = 11.14
    Columns("N:N").ColumnWidth = 10.14
    Columns("M:M").ColumnWidth = 11.14
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "EDİNME " & Chr(10) & "TARİHİ"
    Range("A1:Q4").Select
    Range("Q3").Activate
    With Selection
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    With Selection
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "KİMDEN" & Chr(10) & " ALINDI"
    Range("E4").Select
    Columns("D:D").EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("E4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:Q4").Select
    Range("A4").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Selection.Font.Bold = True
' Format !!**********************************************************************************************************************
'

For AB = 1 To 13
        Do While True
            Sheets(Getthename).Select
            Columns("A:A").Select
            Do While True
                Set FirstCellWithText = Selection.Find(What:=Ay(AB), After:=ActiveCell, LookIn:=xlFormulas, _
                                            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                            MatchCase:=False, SearchFormat:=False)
                If FirstCellWithText Is Nothing Then
                    Exit Do
                Else
                    sd3 = ActiveCell.Row
                    If sd < sd2 Then
                        sd = ActiveCell.Row
                        Selection.FindNext(After:=ActiveCell).Activate
                        sd2 = ActiveCell.Row
                    Else
                        say = say + 1
                        Exit Do
                       
                    End If
                End If
            Loop
            If sd3 = 0 Then
            Exit Do
            Else
                If sd = sd3 Then
                    For A = sd3 To sd
                        Sheets(Getthename).Select
                        Cells(A, 1).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 1).Select
                        ActiveSheet.Paste
                        Columns("A:A").EntireColumn.AutoFit
                    
                        Sheets(Getthename).Select
                        Cells(A, 2).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 2).Select
                        ActiveSheet.Paste
                        Columns("B:B").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 3).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 3).Select
                        ActiveSheet.Paste
                        Columns("C:C").EntireColumn.AutoFit
                    
                        Sheets(Getthename).Select
                        Cells(A, 4).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 4).Select
                        ActiveSheet.Paste
                        Columns("D:D").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 5).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 5).Select
                        ActiveSheet.Paste
                        Columns("E:E").EntireColumn.AutoFit
                    
                        Sheets(Getthename).Select
                        Cells(A, 6).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 6).Select
                        ActiveSheet.Paste
                        Columns("F:F").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 7).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 7).Select
                        ActiveSheet.Paste
                        Columns("G:G").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 8).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 8).Select
                        ActiveSheet.Paste
                        Columns("H:H").EntireColumn.AutoFit
                    
                        Sheets(Getthename).Select
                        Cells(A, 9).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 9).Select
                        ActiveSheet.Paste
                        Columns("I:I").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 10).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 10).Select
                        ActiveSheet.Paste
                        Columns("J:J").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 11).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 11).Select
                        ActiveSheet.Paste
                        Columns("K:K").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 12).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 12).Select
                        ActiveSheet.Paste
                        Columns("L:L").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 13).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 13).Select
                        ActiveSheet.Paste
                        Columns("M:M").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 14).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 14).Select
                        ActiveSheet.Paste
                        Columns("N:N").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 15).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 15).Select
                        ActiveSheet.Paste
                        Columns("O:O").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 16).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 16).Select
                        ActiveSheet.Paste
                        Columns("P:P").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 17).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 17).Select
                        ActiveSheet.Paste
                        Columns("Q:Q").EntireColumn.AutoFit
                    
                        Range(Cells(say, 1), Cells(say, 17)).Select
                        If ABC = 1 Then
                            With Selection.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = Renk(ABC)
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        ABC = 2
                        Else
                            With Selection.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = Renk(ABC)
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        ABC = 1
                        End If
                        say = say + 1
                    Next
                    sd = 0
                    sd3 = 0
                    sd2 = 1
                    Exit Do
                Else
                    For A = sd3 To sd
                        Sheets(Getthename).Select
                        Cells(A, 1).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 1).Select
                        ActiveSheet.Paste
                        Columns("A:A").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 2).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 2).Select
                        ActiveSheet.Paste
                        Columns("B:B").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 3).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 3).Select
                        ActiveSheet.Paste
                        Columns("C:C").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 4).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 4).Select
                        ActiveSheet.Paste
                        Columns("D:D").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 5).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 5).Select
                        ActiveSheet.Paste
                        Columns("E:E").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 6).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 6).Select
                        ActiveSheet.Paste
                        Columns("F:F").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 7).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 7).Select
                        ActiveSheet.Paste
                        Columns("G:G").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 8).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 8).Select
                        ActiveSheet.Paste
                        Columns("H:H").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 9).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 9).Select
                        ActiveSheet.Paste
                        Columns("I:I").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 10).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 10).Select
                        ActiveSheet.Paste
                        Columns("J:J").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 11).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 11).Select
                        ActiveSheet.Paste
                        Columns("K:K").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 12).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 12).Select
                        ActiveSheet.Paste
                        Columns("L:L").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 13).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 13).Select
                        ActiveSheet.Paste
                        Columns("M:M").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 14).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 14).Select
                        ActiveSheet.Paste
                        Columns("N:N").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 15).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 15).Select
                        ActiveSheet.Paste
                        Columns("O:O").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 16).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 16).Select
                        ActiveSheet.Paste
                        Columns("P:P").EntireColumn.AutoFit
                        
                        Sheets(Getthename).Select
                        Cells(A, 17).Select
                        Selection.Copy
                        Sheets("BUSKI").Select
                        Cells(say, 17).Select
                        ActiveSheet.Paste
                        Columns("Q:Q").EntireColumn.AutoFit
                        
                        Range(Cells(say, 1), Cells(say, 17)).Select
                    
                        If ABC = 1 Then
                            With Selection.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = Renk(ABC)
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        ABC = 2
                        Else
                            With Selection.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = Renk(ABC)
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        ABC = 1
                        End If
                        say = say + 1
                    Next
                    sd = 0
                    sd3 = 0
                    sd2 = 1
                    Exit Do
                End If
            End If
        Loop
Next

For AB = 1 To 12

Do While True
            Sheets("BUSKI").Select
            Columns("A:A").Select
            Do While True
                Set FirstCellWithText = Selection.Find(What:=Ay(AB), After:=ActiveCell, LookIn:=xlFormulas, _
                                            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                            MatchCase:=False, SearchFormat:=False)
                If FirstCellWithText Is Nothing Then
                    Exit Do
                Else
                    sd3 = ActiveCell.Row
                    If sd < sd2 Then
                        sd = ActiveCell.Row
                        Selection.FindNext(After:=ActiveCell).Activate
                        sd2 = ActiveCell.Row
                    Else
                        Exit Do
                    End If
                End If
            Loop
  '*******************************************************************************************************************
            
            If sd = 0 Then
                    sd = 0
                    sd3 = 0
                    sd2 = 1
                    Exit Do
            Else
            
            
            
            
'************************************************************************************************************
'***************************************************** Son Kısım ********************************************
                If sd = sd3 Then
                    sd = 0
                    sd3 = 0
                    sd2 = 1
                    Exit Do
                Else
                        ABCD = sd3
                        fark = sd - sd3 + 1
                        
                        For SSD = 0 To fark - 1
                            For UK = 1 To fark - 1
                                If (ABCD + SSD + UK) > sd Then
                                Else
                                
                                If Cells(ABCD + SSD, 5).Value = Cells(ABCD + SSD + UK, 5).Value Then
                                    If Cells(ABCD + SSD, 6).Value = Cells(ABCD + SSD + UK, 6).Value Then
                                        
                                        For ABCDE = 2 To 8
                                        If Range(Cells(ABCD + SSD + UK, ABCDE), Cells(ABCD + SSD + UK, ABCDE)) = "" Then
                                        Else
                                                                                
                                        Range(Cells(ABCD + SSD + UK, ABCDE), Cells(ABCD + SSD + UK, ABCDE)).ClearContents
                                        
                                        Range(Cells(ABCD + SSD, ABCDE), Cells(ABCD + SSD + UK, ABCDE)).Select
                                        Application.CutCopyMode = False
                                        With Selection
                                            .HorizontalAlignment = xlCenter
                                            .VerticalAlignment = xlBottom
                                            .WrapText = False
                                            .Orientation = 0
                                            .AddIndent = True
                                            .IndentLevel = 0
                                            .ShrinkToFit = False
                                            .ReadingOrder = xlContext
                                            .MergeCells = True
                                        End With
                                        Selection.Merge
                                        With Selection
                                            .HorizontalAlignment = xlCenter
                                            .VerticalAlignment = xlCenter
                                            .WrapText = False
                                            .Orientation = 0
                                            .AddIndent = False
                                            .IndentLevel = 0
                                            .ShrinkToFit = False
                                            .ReadingOrder = xlContext
                                            .MergeCells = True
                                        End With
                                        End If
                                     
                                        Next
                                        If Range(Cells(ABCD + SSD + UK, 14), Cells(ABCD + SSD + UK, 14)) = "" Then
                                        Else
                                        Range(Cells(ABCD + SSD + UK, 14), Cells(ABCD + SSD + UK, 14)).ClearContents
                                        Range(Cells(ABCD + SSD, 14), Cells(ABCD + SSD + UK, 14)).Select
                                        Application.CutCopyMode = False
                                        With Selection
                                            .HorizontalAlignment = xlCenter
                                            .VerticalAlignment = xlBottom
                                            .WrapText = False
                                            .Orientation = 0
                                            .AddIndent = True
                                            .IndentLevel = 0
                                            .ShrinkToFit = False
                                            .ReadingOrder = xlContext
                                            .MergeCells = True
                                        End With
                                        Selection.Merge
                                        With Selection
                                            .HorizontalAlignment = xlCenter
                                            .VerticalAlignment = xlCenter
                                            .WrapText = False
                                            .Orientation = 0
                                            .AddIndent = False
                                            .IndentLevel = 0
                                            .ShrinkToFit = False
                                            .ReadingOrder = xlContext
                                            .MergeCells = True
                                        End With
                                        End If
                                        If Range(Cells(ABCD + SSD + UK, 15), Cells(ABCD + SSD + UK, 15)) = "" Then
                                        Else
                                        Range(Cells(ABCD + SSD + UK, 15), Cells(ABCD + SSD + UK, 15)).ClearContents
                                        Range(Cells(ABCD + SSD, 15), Cells(ABCD + SSD + UK, 15)).Select
                                        Application.CutCopyMode = False
                                        With Selection
                                            .HorizontalAlignment = xlCenter
                                            .VerticalAlignment = xlBottom
                                            .WrapText = False
                                            .Orientation = 0
                                            .AddIndent = True
                                            .IndentLevel = 0
                                            .ShrinkToFit = False
                                            .ReadingOrder = xlContext
                                            .MergeCells = True
                                        End With
                                        Selection.Merge
                                        With Selection
                                            .HorizontalAlignment = xlCenter
                                            .VerticalAlignment = xlCenter
                                            .WrapText = False
                                            .Orientation = 0
                                            .AddIndent = False
                                            .IndentLevel = 0
                                            .ShrinkToFit = False
                                            .ReadingOrder = xlContext
                                            .MergeCells = True
                                        End With
                                        End If
                                        If Range(Cells(ABCD + SSD + UK, 17), Cells(ABCD + SSD + UK, 17)) = "" Then
                                        Else
                                        Range(Cells(ABCD + SSD + UK, 17), Cells(ABCD + SSD + UK, 17)).ClearContents
                                        Range(Cells(ABCD + SSD, 17), Cells(ABCD + SSD + UK, 17)).Select
                                        Application.CutCopyMode = False
                                        With Selection
                                            .HorizontalAlignment = xlCenter
                                            .VerticalAlignment = xlBottom
                                            .WrapText = False
                                            .Orientation = 0
                                            .AddIndent = True
                                            .IndentLevel = 0
                                            .ShrinkToFit = False
                                            .ReadingOrder = xlContext
                                            .MergeCells = True
                                        End With
                                        Selection.Merge
                                        With Selection
                                            .HorizontalAlignment = xlCenter
                                            .VerticalAlignment = xlCenter
                                            .WrapText = False
                                            .Orientation = 0
                                            .AddIndent = False
                                            .IndentLevel = 0
                                            .ShrinkToFit = False
                                            .ReadingOrder = xlContext
                                            .MergeCells = True
                                        End With
                                        End If
                                                                        
                                    End If
                                End If
                                End If
                                
                            Next
                         Next
                
                
                
'***************************************************************************************************************
                
                    Range(Cells(sd3 + 1, 1), Cells(sd, 1)).ClearContents
                    Range(Cells(sd3, 1), Cells(sd, 1)).Select
                    Application.CutCopyMode = False
                        With Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlBottom
                            .WrapText = False
                            .Orientation = 0
                            .AddIndent = True
                            .IndentLevel = 0
                            .ShrinkToFit = False
                            .ReadingOrder = xlContext
                            .MergeCells = True
                        End With
                        Selection.Merge
                            With Selection
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                                .WrapText = False
                                .Orientation = 0
                                .AddIndent = False
                                .IndentLevel = 0
                                .ShrinkToFit = False
                                .ReadingOrder = xlContext
                                .MergeCells = True
                            End With
                        sd = 0
                        sd3 = 0
                        sd2 = 1
                        Exit Do
                End If
                
                End If
Loop
Next
sdsem = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
Range(Cells(5, 1), Cells(sdsem, 17)).Select
    ActiveWindow.SmallScroll Down:=-42
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
'****************************************************************** Pivot Tablolar **********************************************
    Dim LastRow As Long
    Dim pvt As PivotTable
    
    LastRow = Range("L" & Rows.Count).End(xlUp).Row
    Range("C4:C" & LastRow).Select
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Grafikler"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "BUSKI!R4C3:R64000C3", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="Grafikler!R3C1", TableName:="PivotTable6", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("Grafikler").Select
    Cells(3, 1).Select
    Set pvt = ActiveSheet.PivotTables("PivotTable6")
    pvt.AddDataField pvt.PivotFields("İLÇE"), "İrtifak Sayısı", xlCount
    ActiveCell.Select
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("İLÇE")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    '**********************************
    Dim pvt2 As PivotTable
    Sheets("BUSKI").Select
    Range("C4:C" & LastRow, "N4:N" & LastRow).Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "BUSKI!R4C3:R64000C12", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="Grafikler!R3C5", TableName:="PivotTable61", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("Grafikler").Select
    Cells(3, 5).Select
    
    Set pvt2 = ActiveSheet.PivotTables("PivotTable61")
    pvt2.AddDataField pvt2.PivotFields("İRTİFAK ALAN" & Chr(10) & "(m2)"), "Toplam İrtifak Alanı", xlSum
    ActiveCell.Select
    
    With ActiveSheet.PivotTables("PivotTable61").PivotFields("İLÇE")
        .Orientation = xlRowField
        .Position = 1
    End With
    '*******************************************
    
    Dim pvt4 As PivotTable
    Sheets("BUSKI").Select
    Range("C4:C" & LastRow, "N4:N" & LastRow).Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "BUSKI!R4C3:R64000C12", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="Grafikler!R3C10", TableName:="PivotTable63", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("Grafikler").Select
    Cells(3, 10).Select
    
    Set pvt4 = ActiveSheet.PivotTables("PivotTable63")
    pvt4.AddDataField pvt4.PivotFields("İRTİFAK DEĞERİ" & Chr(10) & "(TL)"), "Toplam İrtifak Değeri", xlSum
    ActiveCell.Select
    
    With ActiveSheet.PivotTables("PivotTable63").PivotFields("İLÇE")
        .Orientation = xlRowField
        .Position = 1
    End With
       
    
    
    '*******************************************
    Dim pvt3 As PivotTable
    Sheets("BUSKI").Select
    Range("C4:C" & LastRow, "O4:O" & LastRow).Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "BUSKI!R4C3:R64000C15", Version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:="Grafikler!R20C1", TableName:="PivotTable62", DefaultVersion _
        :=xlPivotTableVersion12
    Sheets("Grafikler").Select
    Cells(20, 1).Select
    
    Set pvt3 = ActiveSheet.PivotTables("PivotTable62")
    pvt3.AddDataField pvt3.PivotFields("KULLANIM AMACI"), "KULLANIM AMACI SAYISI", xlCount
    ActiveCell.Select
    
    With ActiveSheet.PivotTables("PivotTable62").PivotFields("KULLANIM AMACI")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    With ActiveSheet.PivotTables("PivotTable62").PivotFields("İLÇE")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    '************************** GRAFİKLER*************************************************************************************************
    
    
    Range("A6").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Grafikler!$A$3:$B$12")
    ActiveChart.Location Where:=xlLocationAsNewSheet
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.ChartArea.Select
    ActiveChart.Legend.Select
    Selection.Delete
    Sheets("Grafikler").Select
    Range("E3").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Grafikler!$E$3:$F$12")
    ActiveChart.Location Where:=xlLocationAsNewSheet
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.ChartArea.Select
    ActiveChart.Legend.Select
    Selection.Delete
    Sheets("Grafikler").Select
    Range("J3").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Grafikler!$J$3:$K$12")
    ActiveChart.Location Where:=xlLocationAsNewSheet
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.ChartArea.Select
    ActiveChart.Legend.Select
    Selection.Delete
    Sheets("Grafikler").Select
    Range("A23").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Grafikler!$A$20:$J$29")
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.ChartArea.Select
    ActiveChart.Location Where:=xlLocationAsNewSheet
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)
    ActiveChart.ChartArea.Select
    ActiveChart.ChartArea.Select
    ActiveChart.Legend.Select
    
End Sub

Function SheetExists(sname, Optional wbName As Variant) As Boolean
    '   check a worksheet exists in the active workbook
    '   or in a passed in optional workbook
        Dim X As Object

        On Error Resume Next
        If IsMissing(wbName) Then
            Set X = ActiveWorkbook.Sheets(sname)
        ElseIf WorkbookIsOpen(wbName) Then
            Set X = Workbooks(wbName).Sheets(sname)
        Else
            SheetExists = False
            Exit Function
        End If

        If Err = 0 Then SheetExists = True _
        Else SheetExists = False
End Function

Function WorkbookIsOpen(wbName) As Boolean
    '   check to see if a workbook is actually open
        Dim X As Workbook
        On Error Resume Next
        Set X = Workbooks(wbName)
        If Err = 0 Then WorkbookIsOpen = True _
        Else WorkbookIsOpen = False
End Function
    
Function WorksheetExists2(WorksheetName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    With wb
        On Error Resume Next
        WorksheetExists2 = (.Sheets(WorksheetName).Name = WorksheetName)
        On Error GoTo 0
    End With
End Function
