Option Explicit
Sub JPPatentDocParagraphNumbering()
'*************************************************************
'* 일본 특허 문서의 Paranum 삽입.
'*
'*************************************************************
  InsertFieldCodeJPParanum
  ReplaceParanumTextswithJPFieldCode
  UpdateAllFieldCode
End Sub
Sub InsertFieldCodeJPParanum()
'*************************************************************
' Insert filed : { SEQ Jpara \#"【0000】" \*DBCHAR }
'
'*************************************************************

  ' 현재 커서 위치에서 줄바꿈해서 새 빈 문단으로 이동한다.
  Selection.TypeParagraph
  ' 필드를 삽입한다.
  Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
  Text:="SEQ Jpara \#" & Chr(34) & ChrW(12304) & "0000" & ChrW(12305) & Chr(34) & " \*DBCHAR", _
    PreserveFormatting:=False
End Sub
Sub ReplaceParanumTextswithJPFieldCode()
'*************************************************************
' InsertFieldCodeJPParanum() 수행 후에 수행해야 한다.
' 영문 문단번호 [0001]^t을 찾아서 필드코드로 바꾼다.
'*************************************************************
  
  Selection.HomeKey Unit:=wdLine, Extend:=True ' 필드코드를 선택하여 복사

  Selection.Copy
  
  ActiveDocument.ConvertNumbersToText wdNumberAllNumbers 'Listnum을 텍스트로 바꿈
  Application.ScreenRefresh
  
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = "\[([0-9]{4,5})\][^t ]@<" ' Paragraph Listnum을 찾기 위한 Wildcard
    .Replacement.Text = "^c^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .CorrectHangulEndings = False
    .HanjaPhoneticHangul = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchFuzzy = False
    .MatchWildcards = True
  End With
  Selection.Find.Execute Replace:=wdReplaceAll
  
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = "\[([0-9]{4,5})\][^t ]@" ' Paragraph Listnum을 찾기 위한 Wildcard
    .Replacement.Text = "^c^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .CorrectHangulEndings = False
    .HanjaPhoneticHangul = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchFuzzy = False
    .MatchWildcards = True
  End With
  Selection.Find.Execute Replace:=wdReplaceAll
  
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = "\[([0-9]{4,5})\]^t" ' Paragraph Listnum을 찾기 위한 Wildcard
    .Replacement.Text = "^c^p"
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchByte = False
    .CorrectHangulEndings = False
    .HanjaPhoneticHangul = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchFuzzy = False
    .MatchWildcards = True
  End With
  Selection.Find.Execute Replace:=wdReplaceAll
  
  
End Sub

Sub UpdateAllFieldCode()
  Selection.WholeStory
  Selection.Fields.Update
  Selection.HomeKey Unit:=wdStory
End Sub
'//////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
