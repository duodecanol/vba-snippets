Option Explicit
Sub DetermineLocation()
  '
  ' DetermineLocation Subprocedure
  ' Main Entry Point

Dim rngDoc As Range
Dim rngSearch As Range
Dim bFound As Boolean
Dim strSearchText As String
Dim strResult As String
Dim strLocation As Object

Dim pageNo As String
Dim paraNo As String
Dim lineNo As String
Dim locus As String
Dim temp As String

Set strLocation = CreateObject("System.Collections.ArrayList") 'https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8

'Dim sssss As ArrayList ' mscorlib.dll REQUIRED
'Set sssss = New ArrayList

Set rngDoc = ActiveDocument.Range
Set rngSearch = rngDoc.Duplicate

strSearchText = InputBox("Enter search word", "search string") ' Trim하지말것. 공백포함해서 Uniqueness 기도하는 경우 존재.

'With ActiveDocument.Content.Find
With rngSearch.Find
  .ClearFormatting
  .Text = strSearchText
  bFound = .Execute(Forward:=True, Wrap:=wdFindContinue)
 
  Do While bFound = True
  
    .Parent.Select
    pageNo = Selection.Information(wdActiveEndPageNumber)  'The page number of the current selected text
    paraNo = Trim(Selection.Paragraphs(1).Range.ListFormat.ListString) 'The paragraph numbering ("[0000]") of the current selected text
    If paraNo = "" Then
      temp = Left(Selection.Paragraphs(1).Range.Text, 8)
      Debug.Print (temp)
      paraNo = Split((RegExpReplace(temp, "\s+", Chr(0))), Chr(0))(0)
      Debug.Print (paraNo)
      
'      Debug.Print (re.Split(temp, "\s")(1))
    End If
    
    If paraNo = "" Then
      
      ' 청구항 번호 얻기/ 문단번호 없는 경우 처리하기.
      lineNo = Selection.Information(wdFirstCharacterLineNumber) 'The line number of the current selected text
      locus = "P" & pageNo & "/" & "L" & lineNo
    Else
      lineNo = GetLineNumberRelativeToParagraph(Selection.Range)
      locus = "P" & pageNo & "/" & paraNo & "/" & "L" & lineNo
    End If
    Debug.Print (locus)
    
    If strLocation.Count > 0 Then
      If strLocation(0) = locus Then
        Debug.Print ("==========================")
        Exit Do 'if N the item is identical to the first item, break loop
      End If
    End If
        
    strLocation.Add locus
        
    rngSearch.Collapse wdCollapseEnd 'search from after the found to the end of the doc
    bFound = .Execute(Forward:=True)
 
  Loop
  
    'If .found = True Then .Parent.Bold = True
End With

Dim item As Variant
For Each item In strLocation
  strResult = strResult + item + ", "
Next item
 strResult = Left(strResult, Len(strResult) - 2)
Debug.Print strResult
MsgBox strResult

Clipboard strResult

End Sub
Function GetLocationInfoSelectionTable(r As Range) As String
  Dim iColNum As Long
  Dim iRowNum As Long
  
  If r.Information(wdWithInTable) <> True Then
    Err.Raise vbObjectError + 999, "GetLocationInfoSelectionTable", "Selection range is not in a table."
    Exit Function
  End If
  
  iColNum = r.Information(wdStartOfRangeColumnNumber)
  iRowNum = r.Information(wdStartOfRangeRowNumber) ' 2페이지 이상의 표에서 Row값을 구하는 것은 추후 구현.
  
  GetLocationInfoSelectionTable = "Col " & iColNum & "/Row " & iRowNum
End Function
Function GetParaNum(r As Range) As Long
  Dim rParagraphs As Range
    
  Set rParagraphs = ActiveDocument.Range(Start:=0, End:=r.End)
  GetParaNum = rParagraphs.Paragraphs.Count
End Function


'//////////////////////////////////////////////////////////
'// 1. 한페이지 내
'// 2. 다음페이지에 걸친 문단
'// 3. 3페이지 이상에 걸친 문단
'//////////////////////////////////////////////////////////

Function GetLineNumberRelativeToParagraph(r As Range) As Long
  Dim iSelectionParaNum As Long
  Dim iLineNumberRelativeToPage As Long
  Dim rngSelectionRangeFromParaFirstCharacter As Range
  
  iSelectionParaNum = GetParaNum(Selection.Range)
  iLineNumberRelativeToPage = r.Information(wdFirstCharacterLineNumber)
  Set rngSelectionRangeFromParaFirstCharacter = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(iSelectionParaNum).Range.Characters(1).End, _
  End:=r.End)

  GetLineNumberRelativeToParagraph = rngSelectionRangeFromParaFirstCharacter.ComputeStatistics(wdStatisticLines)
End Function

Sub PrintLineNumberRelativeToParagraph() ' For test
  If Selection.Range.ComputeStatistics(wdStatisticLines) > 1 Then
    Debug.Print "L " & GetLineNumberRelativeToParagraph(Selection.Characters(1)) & _
                " ~ " & GetLineNumberRelativeToParagraph(Selection.Characters(Selection.Characters.Count))
  Else
    Debug.Print "L  " & GetLineNumberRelativeToParagraph(Selection.Range)
  End If
End Sub

Sub PrintLocationInfoSelectionTable()
  Debug.Print GetLocationInfoSelectionTable(Selection.Range)
End Sub

Sub testme()
  SplitArr = Split((RegExpReplace("my+test*exp?split.pattern", "[+*?.]", Chr(0))), Chr(0))
  'Debug.Print (SplitArr)
    
End Sub

Function RegExpReplace(ByVal WhichString As String, _
  ByVal Pattern As String, _
  ByVal ReplaceWith As String, _
  Optional ByVal IsGlobal As Boolean = True, _
  Optional ByVal IsCaseSensitive As Boolean = True) As String
    
Dim objRegExp As Object
Set objRegExp = CreateObject("vbscript.regexp")
    objRegExp.Global = IsGlobal
    objRegExp.Pattern = Pattern
    objRegExp.IgnoreCase = Not IsCaseSensitive
    RegExpReplace = objRegExp.Replace(WhichString, ReplaceWith)
End Function

