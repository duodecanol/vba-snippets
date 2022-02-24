Option Explicit
Sub bilingualColumnFormatEdit()
'
' bilingualColumnFormatEdit Macro
' for two column translation request
'
Application.ScreenUpdating = False

'Delete Head Row
  ActiveDocument.Tables(1).Rows(1).Select
  Selection.Rows.Delete
  
' Delete Status column
  ActiveDocument.Tables(1).Columns(5).Select
  Selection.Cells.Delete ShiftCells:=wdDeleteCellsEntireColumn
  
' Remove segment code
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = _
      " ^l[a-z0-9]{8}\-[a-z0-9]{4}\-[a-z0-9]{4}\-[a-z0-9]{4}\-[a-z0-9]{12}"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .CorrectHangulEndings = True
    .HanjaPhoneticHangul = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchFuzzy = False
    .MatchWildcards = True
  End With
  
  Selection.Find.Execute
  Selection.Find.Execute Replace:=wdReplaceAll
' widen comment column
  Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=145.05, RulerStyle:=wdAdjustNone
  
Application.ScreenUpdating = True

End Sub

Sub trilingualColumnFormatEdit()
'
' trilingualColumnFormatEdit Macro
' for three column translation request for 4DIP
'

' Page size to A3
  With Selection.PageSetup
    .LineNumbering.Active = False
    .PageWidth = CentimetersToPoints(41.99)
    .PageHeight = CentimetersToPoints(29.7)
    .LayoutMode = wdLayoutModeDefault
  End With
  
' Delete Status column
  ActiveDocument.Tables(1).Columns(6).Select
  Selection.Cells.Delete ShiftCells:=wdDeleteCellsEntireColumn
  
' Remove segment code
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = _
      " ^l[a-z0-9]{8}\-[a-z0-9]{4}\-[a-z0-9]{4}\-[a-z0-9]{4}\-[a-z0-9]{12}"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .CorrectHangulEndings = True
    .HanjaPhoneticHangul = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchFuzzy = False
    .MatchWildcards = True
  End With
  Selection.Find.Execute
  Selection.Find.Execute Replace:=wdReplaceAll
  
' widen comment column
  Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=192.4, RulerStyle:=wdAdjustNone
  
' Reset find & replace conditions
  ResetFRcond
End Sub

Sub highlightPretranslatedRows()
  Dim table As table
  Dim rowNum As Integer
  Dim i As Integer
  Dim cellText As String
  
  'Application.ScreenUpdating = False
  
  Set table = ActiveDocument.Tables(1)
  rowNum = table.Rows.Count
  
  For i = 1 To rowNum
    table.Cell(i, 3).Select
    cellText = table.Cell(i, 3).Range.Text
    ' Debug.Print (cellText)
    
    ' Filtering pretranslated segments. empty segs in Japnanese Col except for 식별항목
    If Len(cellText) > 3 And Left(cellText, 1) <> ChrW(12304) Then
      table.Rows(i).Select
      Selection.Shading.Texture = wdTextureNone
      Selection.Shading.ForegroundPatternColor = wdColorAutomatic
      Selection.Shading.BackgroundPatternColor = -671023207
    End If
  Next i
  
  ActiveDocument.Range(0, 0).Select
  Call insertDoNotTranslateNotification
  
  'Application.ScreenUpdating = True

End Sub

Sub highlightPretranslatedRowsBackslash()
  Dim table As table
  Dim rowNum As Integer
  Dim i As Integer
  Dim cellText As String
  
  Set table = ActiveDocument.Tables(1)
  rowNum = table.Rows.Count
  
  For i = 1 To rowNum
    table.Cell(i, 3).Select
    cellText = table.Cell(i, 3).Range.Text
    ' Debug.Print (cellText)
    
    ' Filtering pretranslated segments. empty segs in Japnanese Col except for 식별항목
    If Left(cellText, 1) = "\" Then
      table.Rows(i).Select
      Selection.Shading.Texture = wdTextureNone
      Selection.Shading.ForegroundPatternColor = wdColorAutomatic
      Selection.Shading.BackgroundPatternColor = -671023207
    End If
  Next i
  
  ActiveDocument.Range(0, 0).Select

End Sub
Sub insertDoNotTranslateNotificationKR()
  Dim table As table
  
  Set table = ActiveDocument.Tables(1)
  table.Select
  Selection.SplitTable
  
  With ActiveDocument.Range(0, 0)
   .InsertParagraph
   .InsertBefore "푸른색으로 하이라이팅된 행들은 번역분량에서 제외되므로 번역이 불필요합니다. 검토 또한 불필요하며, 추후 당사 처리합니다."
   .Font.Name = "맑은 고딕"
   .Font.Size = 18
   .Font.ColorIndex = wdRed
   .ParagraphFormat.LeftIndent = InchesToPoints(3)
   .ParagraphFormat.RightIndent = InchesToPoints(3)
   .ParagraphFormat.Alignment = wdAlignParagraphCenter
  End With
End Sub
