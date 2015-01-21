' Run the following routine then copy and paste the results into HTML Tidy at
' http://infohound.net/tidy/tidy.pl
' Then copy the results into ERM
'
Sub convert2html()
' Converts a Microsoft Word document into simplified html code.
'
' Alternative way of saving Word documents as webpages.  This scripts converts the .doc(x) straight into html devoid of
' any MS Word styles, etc.
'
' By Rob Vance

    ActiveDocument.Save                                 'Save current document
    Application.ScreenUpdating = False                  'screen updates off during macro

    'replace_headers                                      'replace headers
    first_page                                          'covert page
    TableHeadersRepeat                                  'repeat table headers
    replace_properties                                  'replace properties
    remove_headers                                      'remove all of the header information
    remove_footers                                      'remove all of the footer information
    remove_toc                                          'table of contents
    'delete_string ("Table of Contents")
    insert_pagebreak ("Document Management")
    
    replace_hdrs                                        'convert custom styles
    replace_format_string ("Document Management")
    replace_format_string ("Document Change Log")
    replace_format_string ("Stakeholder Review")
    
    replace_notes                                       'replace footnotes and endnotes

    replace_tables                                      'convert from tables
    
    replace_lists                                       'convert simple one level lists
    replace_hyper                                       'convert hyperlinks
    
    replace_formated_paragraphs                         'convert sections with formatting
    
    'replace_other_paragraphs                             'convert other sections
    place_headerfooter                                   'move HTML header
    saveashtml                                          'save changes
    
    ActiveDocument.Select
    Selection.ClearFormatting
    Application.ScreenUpdating = True
End Sub
Function TableHeadersRepeat()
Dim objTable As Word.Table
Dim myRow As Row
Dim pge As Page
Dim pg As Long

For pg = 1 To ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
'If pg = 1 Then
'    Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext
'Else
If pg <> 1 Then
    For Each objTable In ActiveDocument.Tables
    On Error Resume Next
        objTable.Rows(1).HeadingFormat = True
        objTable.Rows.AllowBreakAcrossPages = True
        For Each myRow In objTable.Rows
        On Error Resume Next
        myRow.HeightRule = wdRowHeightAuto
        myRow.Cells.VerticalAlignment = wdCellAlignVerticalTop
        If myRow.Cells(1).RowIndex = 1 Then
            'myRow.Range.Style = "TBL Header"
            'If ActiveDocument.Range(0, Selection.Tables(1).Range.End).Tables.Count <> 1 Then        ' ignore the first table i.e., title page
                myRow.Range.Shading.BackgroundPatternColor = RGB(191, 191, 191)
                myRow.HeadingFormat = True
            'End If
        End If
        Next myRow
    Next objTable
End If
Next pg
Set objTable = Nothing
End Function
Function replace_properties()
'Get the Title property
    oFN = ActiveDocument.Name
    oFN = Left(oFN, Len(oFN) - 5)
    oFN = Right(oFN, Len(oFN) - 6)
    sFN = Split(oFN, " - ")
    sTitle = ActiveDocument.BuiltInDocumentProperties("Title").Value
    
' Update Properties
    ActiveDocument.BuiltInDocumentProperties("Author") = "Robert Vance"
    ActiveDocument.BuiltInDocumentProperties("Manager") = ""
    ActiveDocument.BuiltInDocumentProperties("Company") = "Surescripts, LLC"
    ActiveDocument.BuiltInDocumentProperties("Category") = ""
    ActiveDocument.BuiltInDocumentProperties("Keywords") = sTitle
    ActiveDocument.BuiltInDocumentProperties("Comments") = ""
    mySTR = UpCustoms("CISO", "Director of Security and/or their designee")
    mySTR = UpCustoms("Status", "FINAL")
    mySTR = UpCustoms("Copyright", "Copyright 2014 Surescripts, LLC")
    mySTR = UpCustoms("ISODomain0", "")
    mySTR = UpCustoms("WorkflowChangePath", "")

' Find and Replace text
With ActiveDocument.Content.Find
    .Forward = True
    .Wrap = wdFindStop
    .Execute FindText:="[To be replaced]", ReplaceWith:="[Replace with]", Replace:=wdReplaceAll
End With
End Function
Function UpCustoms(strName, strValue)
On Error Resume Next
If ActiveDocument.CustomDocumentProperties.Count > 0 Then
    For i = 1 To ActiveDocument.CustomDocumentProperties.Count
        ActiveDocument.CustomDocumentProperties(i).Delete
        If ActiveDocument.CustomDocumentProperties(i).Name = strName Then
            ActiveDocument.CustomDocumentProperties(strName) = strValue
        Else
            ActiveDocument.CustomDocumentProperties(strName).Delete
            ActiveDocument.CustomDocumentProperties.Add Name:=strName, LinkToContent:=False, Type:=msoPropertyTypeString, Value:=strValue
        End If
    Next
End If
End Function
Function replace_headers()
Dim i, headnum As Integer

For i = -2 To -7 Step -1
headcount = Abs(i) - 1
    Selection.HomeKey wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles(i)
    Selection.Find.Text = ""
        Do While Selection.Find.Execute = True
            If Selection.Characters.Count > 1 Then
                While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
                    Selection.MoveEnd Unit:=wdCharacter, Count:=-1
                Wend
               Selection.InsertBefore "<h" & headnum & ">"
               Selection.InsertAfter "</h" & headnum & ">"
               Selection.Find.Replacement.ClearFormatting
            End If
        Selection.Style = -1
        Selection.MoveRight wdCharacter, 1
        Loop
Next i

End Function
Function remove_toc()
Dim doc As Document
Dim fld As Field
Dim rng As Range
Dim TOC As TableOfContents
Set doc = ActiveDocument
For Each fld In doc.Fields
    If fld.Type = wdFieldTOC Then
        'fld.Select
        'Selection.Collapse
        'Set rng = Selection.Range 'capture place to re-insert TOC later
        'fld.Cut
        ActiveDocument.TablesOfContents(1).Delete
    End If
Next
On Error Resume Next
If ActiveDocument.Styles("TOC Title").NameLocal = "TOC Title" Then
    ActiveDocument.Styles("TOC Title").Delete
End If

End Function
Function remove_headers()
     Dim sec As Section
     Dim hdr As HeaderFooter
     For Each sec In ActiveDocument.Sections
         For Each hdr In sec.Headers
             If hdr.Exists Then
                 With hdr.Range
                    .Delete
                    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                    '.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                    '.Borders(wdBorderRight).LineStyle = wdLineStyleNone
                End With
             End If
         Next hdr
     Next sec
End Function
Function remove_footers()
     Dim sec As Section
     Dim ftr As HeaderFooter
     For Each sec In ActiveDocument.Sections
         For Each ftr In sec.Footers
             If ftr.Exists Then
                 With ftr.Range
                    .Delete
                    .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                    '.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                    .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                    '.Borders(wdBorderRight).LineStyle = wdLineStyleNone
                End With
             End If
         Next ftr
     Next sec
End Function
Function replace_formating()
'This macro searches the text to any text in italic, bold or underlined and converts it into HTML
Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.Font.Italic = True
Selection.Find.Text = ""
Do While Selection.Find.Execute = True
    Selection.Font.Italic = False
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    Wend
    Selection.InsertBefore "<i>"
    Selection.InsertAfter "</i>"
    
    Selection.MoveRight wdCharacter, 1
Loop

Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.Font.Bold = True
Selection.Find.Text = ""
Do While Selection.Find.Execute = True
    Selection.Font.Bold = False
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    Wend
    Selection.InsertBefore "<b>"
    Selection.InsertAfter "</b>"
    Selection.MoveRight wdCharacter, 1
Loop

Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.Font.Underline = True
Selection.Find.Text = ""
Do While Selection.Find.Execute = True
    Selection.Font.Underline = False
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    Wend
    Selection.InsertBefore "<u>"
    Selection.InsertAfter "</u>"
    Selection.MoveRight wdCharacter, 1
Loop

End Function
Function replace_notes()
' Footnotes convert into endnotes
Dim num As Long
Dim myString As String
Dim myStoryRange As Range

With ActiveDocument.Sections.Last.Range
    .Collapse Direction:=wdCollapseEnd
    .InsertParagraphAfter
If ActiveDocument.Endnotes.Count = 0 Then
    Exit Function
Else
    .InsertAfter "<hr />" & vbCr
End If
    Selection.EndKey Unit:=wdStory
    Selection.ClearFormatting
End With

If ActiveDocument.Footnotes.Count > 0 Then
    ActiveDocument.Footnotes.Convert
End If
        
If ActiveDocument.Endnotes.Count = 0 Then
    Exit Function
End If

With Selection
    .HomeKey wdStory
    For num = 1 To ActiveDocument.Endnotes.Count
        .GoToNext wdGoToEndnote
        .TypeText Text:="<a href=" & Chr(34) & "#edn" & CStr(num) & Chr(34) & " name=" & Chr(34) & "endref" & CStr(num) & Chr(34) & "><sup>" & CStr(num) & "</sup></a>"
        .Expand wdWord
        With ActiveDocument.Endnotes(1)
            myString = myString & "<a href=" & Chr(34) & "#ednref" & CStr(num) & Chr(34) & " name=" & Chr(34) & "end" & CStr(num) & Chr(34) & "><sup>" & CStr(num) & "</sup></a>" & "<span style='font-family: Arial; font-size: 8px;'>" & .Range.Text & "</span>" & vbCrLf
            .Delete
        End With
    Next
    .EndKey wdStory
    .InsertAfter myString
    .Collapse Direction:=wdCollapseEnd
End With

End Function
Function replace_lists()
Dim lstVAL As String, lstLVL As String, outLVL As String
Dim sText As String
Dim lijst As List
Dim para As Paragraph

'replace_hdrs

Selection.HomeKey wdStory

For Each para In ActiveDocument.ListParagraphs
    lstVAL = para.Range.ListFormat.ListValue
    lstLVL = para.Range.ListFormat.ListLevelNumber
    lstTYP = para.Range.ListFormat.ListType
    para.Range.Select
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    Wend
    'sText = Selection.Text
    'MsgBox "Paragraph type=" & lstTYP & " Level " & lstLVL & vbCr & vbCr & sText
    If lstTYP <> 4 Then
        Selection.MoveEnd Unit:=wdCharacter, Count:=0
        Selection.InsertBefore "<li><span style='font-family: Arial; font-size: 11px;'>"
        Selection.InsertAfter "</span></li>"
    End If
Next para

For Each lijst In ActiveDocument.Lists
    lstVAL = lijst.Range.ListFormat.ListValue
    lstLVL = lijst.Range.ListFormat.ListLevelNumber
    lstTYP = lijst.Range.ListFormat.ListType
    lijst.Range.Select
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    Wend
    sText = Selection.Text

    If lstTYP <> 4 Then
        If lstTYP = 2 Then
            If lstLVL = 1 Then
                'MsgBox "Lists type=" & lstTYP & " Level " & lstLVL & vbCr & vbCr & sText
                Selection.InsertBefore "<ul type='disc'>"
                Selection.InsertAfter "</ul>"
            Else
                'MsgBox "Lists type=" & lstTYP & " Level " & lstLVL & vbCr & vbCr & sText
                Selection.InsertBefore "<ul type='disc'>"
                Selection.InsertAfter "</ul>"
            End If
        End If
        If lstTYP = 5 Then
            If lstLVL = 1 Then
                'MsgBox "Lists type=" & lstTYP & " Level " & lstLVL & vbCr & vbCr & sText
                Selection.InsertBefore "<ul>"
                Selection.InsertAfter "</ul>"
            End If
        End If
    End If
Next lijst

End Function
Function ToBulletOrNotToBullet()
     Dim para As Paragraph, i As Long
     For Each para In ActiveDocument.Paragraphs
         i = i + 1
         If para.Range.ListFormat.ListType = wdListBullet Then
             MsgBox "Paragraph " & i & " is in a bulleted list and is at Level " _
                 & para.Range.ListFormat.ListLevelNumber
         End If
     Next para
End Function

Private Sub replace_tables()
' convert tables
rtnTables:
On Error GoTo ErrorHandler
Dim oRow As Row
Dim oCell As Cell
Dim sCellText As String
Dim tTable As Table
Dim noRows, noCells, tblnum As Long

        For Each tTable In ActiveDocument.Tables
        On Error GoTo ErrorHandler
            For Each oRow In tTable.Rows
            
                For Each oCell In oRow.Cells
                    sCellText = oCell.Range
                    sCellText = Left$(sCellText, Len(sCellText) - 2)
                    If Len(sCellText) = 0 Then sCellText = "&nbsp;"
                    If oRow.Cells(1).RowIndex = 1 Then
                        sCellText = "<td style='border: 1px solid rgb(0, 0, 0);border-collapse:collapse;'><span style='font-family: Arial; font-size: 11px;'><strong>" & sCellText & "</strong></span></td>"
                    Else
                        sCellText = "<td style='border: 1px solid rgb(0, 0, 0);border-collapse:collapse;'><span style='font-family: Arial; font-size: 11px;'>" & sCellText & "</span></td>"
                    End If
                    oCell.Range = sCellText
                Next oCell
                sCellText = oRow.Cells(1).Range
                sCellText = Left$(sCellText, Len(sCellText) - 2)
                If oRow.Cells(1).RowIndex = 1 Then
                    sCellText = "<tr bgcolor=rgb(191,191,191)>" & vbCr & sCellText
                Else
                    sCellText = "<tr bgcolor='white'>" & vbCr & sCellText
                End If
                oRow.Cells(1).Range = sCellText
                sCellText = oRow.Cells(oRow.Cells.Count).Range
                sCellText = Left$(sCellText, Len(sCellText) - 2)
                sCellText = sCellText & vbCr & "</tr>"
                oRow.Cells(oRow.Cells.Count).Range = sCellText
            Next oRow
            sCellText = tTable.Rows(1).Cells(1).Range
            sCellText = Left$(sCellText, Len(sCellText) - 2)
            sCellText = "<table border='1' style='background-color:#FFFFFF;border-collapse:collapse;border:1px solid #000000;color:#000000;width:75%' cellpadding='1' cellspacing='1'>" & vbCr & sCellText
            tTable.Rows(1).Cells(1).Range = sCellText
            noRows = tTable.Rows.Count
            noCells = tTable.Rows(noRows).Cells.Count
            sCellText = tTable.Rows(noRows).Cells(noCells).Range
            sCellText = Left$(sCellText, Len(sCellText) - 2)
            sCellText = sCellText & vbCr & "</table>"
            tTable.Rows(noRows).Cells(noCells).Range = sCellText
            tTable.ConvertToText Separator:=wdSeparateByParagraphs
        Next tTable
        
ErrorHandler:
    If Err.Number = 5991 Then
        tblnum = ActiveDocument.Range(0, Selection.Sections(1).Range.End).Sections.Count
        MsgBox "Table #" & tblnum
        splitTBL (tblnum)
        GoTo rtnTables
    End If
End Sub
Function splitTBL(ByVal tblnum As Long)
         For Each c In ActiveDocument.Tables(tblnum).Range.Cells
         On Error Resume Next
             c.Select
             RowSpan = (Selection.Information(wdEndOfRangeRowNumber) - Selection.Information(wdStartOfRangeRowNumber)) + 1
             If RowSpan > 1 Then
                 iCol = Selection.Information(wdEndOfRangeColumnNumber)
                 iRow = Selection.Information(wdEndOfRangeRowNumber)
                 'MsgBox "Row: " & iRow & vbCr & "Column: " & iCol
                 'MsgBox "Manually split the merged rows before running the script."
                 Selection.Cells.Split NumRows:=RowSpan, NumColumns:=1, MergeBeforeSplit:=False
             End If
         Next c
End Function
Function replace_hyper()
'convert hyperlinks
Dim hyperCount, i As Long
Dim addr As String

hyperCount = ActiveDocument.Hyperlinks.Count
If hyperCount > 0 Then
    For i = 1 To hyperCount
        With ActiveDocument.Hyperlinks(1)
            addr = .Address
            .Delete
            .Range.InsertBefore "<a href=" & Chr(34) & addr & Chr(34) & ">"
            .Range.InsertAfter "</a>"
        End With
    Next i
End If

End Function
Function replace_pics()
Dim sDir
Dim iDir, num As Integer
Dim oPicture As Word.InlineShape                                    ' Word Shape Object
Dim CurrentMap, ExportMap As String
Dim imgname, oldname As String

CurrentMap = ActiveDocument.Path                                    'Directory from current file
ExportMap = CurrentMap & "\Save_As_HTML_files\"                     'Directory where the images are stored temporarily

On Error Resume Next
Kill CurrentMap & "\Save_As_HTML.html"

On Error Resume Next
Kill ExportMap & "*.*"

On Error Resume Next
RmDir ExportMap
    
Application.Documents.Add ActiveDocument.FullName
ActiveDocument.SaveAs CurrentMap & "\Save_As_HTML.html", FileFormat:=wdFormatHTML
ActiveDocument.Close

' <============= Fixup this section =============>
' Fix where the location of the images are stored within Keylight
num = 1
For Each oPicture In ActiveDocument.InlineShapes
   With oPicture.Range
       imgname = "image" & Format(num, "000") & ".jpg"
       oldname = ExportMap & "image" & Format(num, "000") & ".jpg"
       .InsertBefore "<img src=" & Chr(34) & imgname & Chr(34) & " />"
       oPicture.Delete
       imgname = CurrentMap & "\" & imgname
       FileCopy oldname, imgname
       num = num + 1
   End With
Next

On Error Resume Next
Kill CurrentMap & "\Save_As_HTML.html"

On Error Resume Next
Kill ExportMap & "*.*"

On Error Resume Next
RmDir ExportMap

End Function

Function replace_customparagraphs()
Dim answer, strIn As String

answer = vbYes

Do While answer <> vbNo
    answer = MsgBox("Other paragraph styles convert?", vbQuestion + vbYesNo, "Sections")
    If answer = vbNo Then Exit Function

    strIn = InputBox("Which Style?")

    Selection.HomeKey wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles(strIn)
    Selection.Find.Text = ""
    Do While Selection.Find.Execute = True
        While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
            Selection.MoveEnd Unit:=wdCharacter, Count:=-1
        Wend
    Selection.InsertBefore "<p class=" & Chr(34) & strIn & Chr(34) & ">"
    Selection.InsertAfter "</p>"
    Selection.Find.Replacement.ClearFormatting
    Selection.Style = -1
    Selection.MoveRight wdCharacter, 1
    Loop
Loop

End Function
Function replace_formated_paragraphs()
'sections that are not headers, bullets

Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.MatchWildcards = True
Selection.Find.Text = "^13[a-zA-Z]*^13"
Do While Selection.Find.Execute = True
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    Wend
    Selection.MoveStart Unit:=wdCharacter, Count:=1
    If Selection.Paragraphs.LeftIndent > 60 Then
        'MsgBox Selection.Paragraphs.LeftIndent
        Selection.InsertBefore "<p style='margin-left: 40px;'><span style='font-family: Arial; font-size: 11px;'>"
    Else
        Selection.InsertBefore "<p><span style='font-family: Arial; font-size: 11px;'>"
    End If
    Selection.InsertAfter "</span></p>"
    Selection.Style = -1
    Selection.MoveRight wdCharacter, 1
Loop
End Function
Function replace_format_string(strIn As String)
'sections that are not headers, bullets
Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.MatchWildcards = True
Selection.Find.Text = strIn
Do While Selection.Find.Execute = True
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=0
    Wend
    Selection.MoveStart Unit:=wdCharacter, Count:=0
    Selection.InsertBefore "<p><strong><span style='font-family: Arial; font-size: 14px;'>"
    Selection.InsertAfter "</strong></p>"
    Selection.Style = -1
    Selection.MoveRight wdCharacter, 1
Loop

End Function
Function delete_string(strIn As String)
'sections that are not headers, bullets
Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.MatchWildcards = True
Selection.Find.Text = strIn
Do While Selection.Find.Execute = True
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=0
    Wend
    Selection.MoveStart Unit:=wdCharacter, Count:=0
    Selection.Delete
    Selection.Style = -1
    Selection.MoveRight wdCharacter, 1
Loop
End Function
Function insert_pagebreak(strIn As String)
'sections that are not headers, bullets
Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.MatchWildcards = True
Selection.Find.Text = strIn
Do While Selection.Find.Execute = True
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=0
    Wend
    Selection.MoveStart Unit:=wdCharacter, Count:=0
    Selection.InsertBefore "<br style='clear: both; page-break-before: always;'>" & vbCr
    Selection.Style = -1
    Selection.MoveRight wdCharacter, 1
Loop
End Function
Function replace_empty_paragraphs()
'replace paragraphs empty

Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.MatchWildcards = True
Selection.Find.Text = "^13^13"
Do While Selection.Find.Execute = True
    Selection.MoveStart Unit:=wdCharacter, Count:=1
        Selection.InsertBefore "<p>&nbsp;</p>"
    Selection.Style = -1
    Selection.MoveRight wdCharacter, 1
Loop

End Function
Function replace_other_paragraphs()

Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.MatchWildcards = True
Selection.Find.Text = "^13[A-Z]*^13"
Do While Selection.Find.Execute = True
    While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    Wend
    Selection.MoveStart Unit:=wdCharacter, Count:=1
    Selection.InsertBefore "<p>"
    Selection.InsertAfter "</p>"
    Selection.Style = -1
    Selection.MoveRight wdCharacter, 1
Loop
End Function

Function place_headerfooter()
Dim MyText, strIn As String
Dim MyRange As Object

Set MyRange = ActiveDocument.Range

MyText = "<!DOCTYPE html PUBLIC" & Chr(34) & "-//W3C//DTD XHTML 1.0 Transitional//EN" & Chr(34) & vbCr & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" & Chr(34) & ">" & vbCr & "<html>" & vbCr & "<head>" & vbCr & "<title></title>" & vbCr & "</head>" & vbCr & "<body>" & vbCr
MyRange.InsertBefore (MyText)
MyText = "</body>" & vbCr & "</html>"
MyRange.InsertAfter (MyText)

End Function
Function saveashtml()
Dim filesaveas, answer As String
Dim extPos As Integer
answer = MsgBox("HTML Save?", vbQuestion + vbYesNo, "Save")
If answer = vbNo Then Exit Function
extPos = InStrRev(ActiveDocument.FullName, ".") 'the last point of the search extension. This may namely 3 or 4,
filesaveas = Left(ActiveDocument.FullName, extPos - 1) & ".html"
ActiveDocument.SaveEncoding = msoEncodingUTF8
ActiveDocument.SaveAs Filename:=filesaveas, FileFormat:=wdFormatText
Call tidy(filesaveas)
MsgBox "File saved as " & filesaveas, vbInformation + vbOKOnly, "Ready!"
End Function
Function RemoveAllComments()
    Dim n As Long
    For n = ActiveDocument.Comments.Count To 1 Step -1
    ActiveDocument.Comments(n).Delete
    Next 'n
End Function
Private Sub first_page()
' tables
Dim oRow As Row
Dim oCell As Cell
Dim sCellText As String
' pages
Dim pge As Page
Dim pg As Long
pg = 1

For Each pge In ActiveDocument.ActiveWindow.Panes(1).Pages
    If (pg = 1) Then
        For Each tTable In ActiveDocument.Tables
            tTable.Select
            Selection.ClearFormatting
            For Each oRow In tTable.Rows
                For Each oCell In oRow.Cells
                    sCellText = oCell.Range
                    If (oRow.Cells(1).RowIndex = 1) Then
                        sCellText = "<td style='border: 1px solid rgb(255,255,255);border-collapse:collapse;'><img style='width: 617px; height: 176px;' src='/UI/ImageViewer.aspx?id=788'></td>"
                        oCell.Range = sCellText
                    Else
                        sCellText = Left$(sCellText, Len(sCellText) - 2)
                        sCellText = StrConv(sCellText, vbProperCase)
                        If FindString(LCase(sCellText), LCase("Date")) Then
                            sCellText = Replace(sCellText, "Date", "</strong></td></tr><tr bgcolor='white'><td style='border: 1px solid rgb(255,255,255);border-collapse:collapse;'><strong><span style='font-family: Arial;font-size: 18px;'>Date")
                        End If
                        If Len(sCellText) = 0 Then sCellText = "&nbsp;"
                        sCellText = "<td style='border: 1px solid rgb(255,255,255);border-collapse:collapse;'><strong><span style='font-family: Arial;font-size: 18px;'>" & sCellText & "</span></strong></td>"
                        oCell.Range = sCellText
                    End If
                Next oCell
                sCellText = oRow.Cells(1).Range
                sCellText = Left$(sCellText, Len(sCellText) - 2)
                If sCellText = "Date" Then MsgBox "Found"
                sCellText = "<tr bgcolor='white'>" & vbCr & sCellText
                oRow.Cells(1).Range = sCellText
                sCellText = oRow.Cells(oRow.Cells.Count).Range
                sCellText = Left$(sCellText, Len(sCellText) - 1)
                sCellText = sCellText & vbCr & "</tr>"
                oRow.Cells(oRow.Cells.Count).Range = sCellText
            Next oRow
        sCellText = tTable.Rows(1).Cells(1).Range
        sCellText = Left$(sCellText, Len(sCellText) - 2)
        sCellText = "<table border='1' style='background-color:#FFFFFF;border-collapse:collapse;border:1px solid #FFFFFF;color:#000000;width:50%' cellpadding='1' cellspacing='1'>" & vbCr & sCellText
        tTable.Rows(1).Cells(1).Range = sCellText
        
        tTable.Rows.Add
        sCellText = "<tr bgcolor='white'>" & vbCr & "<tr bgcolor='white'><td style='border: 1px solid rgb(255,255,255);border-collapse:collapse;'><strong><span style='font-family: Arial;font-size: 18px;'>Proprietary and Confidential</td></tr>"
        tTable.Rows(tTable.Rows.Count).Range = sCellText
        
        noRows = tTable.Rows.Count
        noCells = tTable.Rows(noRows).Cells.Count
        sCellText = tTable.Rows(noRows).Cells(noCells).Range
        sCellText = Left$(sCellText, Len(sCellText) - 2)
        sCellText = sCellText & vbCr & "</table>"
        tTable.Rows(noRows).Cells(noCells).Range = sCellText
        tTable.ConvertToText Separator:=wdSeparateByParagraphs
        If (pg = 1) Then Exit For                                                ' the secret sauce
    Next tTable
    End If
    pg = pg + 1
Next pge
End Sub
Function FindString(strCheck As String, strFind As String) As Boolean
    intPos = 0
    intPos = InStr(strCheck, strFind)
    FindString = intPos > 0
End Function
Function cellSel()
'To select a range of cells within a table, declare a Range variable, assign to it the cells you want to select, and then select the range
    Dim myCells As Range
    With ActiveDocument
        Set myCells = .Range(start:=.Tables(1).Cell(1, 1).Range.start, _
            End:=.Tables(1).Cell(1, 4).Range.End)
        myCells.Select
    End With
End Function
Function AddPics()
'http://www.vbaexpress.com/forum/showthread.php?44473-Insert-Multiple-Pictures-Into-Table-Word-With-Macro/page2
    Application.ScreenUpdating = False
    Dim oTbl As Table, i As Long, j As Long, k As Long, strTxt As String
     'Select and insert the Pics
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select image files and click OK"
        .Filters.Add "Images", "*.gif; *.jpg; *.jpeg; *.bmp; *.tif; *.png"
        .FilterIndex = 2
        If .Show = -1 Then
             'Add a 2-row by 2-column table with 7cm columns to take the images
            Set oTbl = Selection.Tables.Add(Selection.Range, 2, 2)
            With oTbl
                .AutoFitBehavior (wdAutoFitFixed)
                .Columns.Width = CentimetersToPoints(7)
                 'Format the rows
                Call FormatRows(oTbl, 1)
            End With
            For i = 1 To .SelectedItems.Count
                j = Int((i + 1) / 2) * 2 - 1
                k = (i - 1) Mod 2 + 1
                 'Add extra rows as needed
                If j > oTbl.Rows.Count Then
                    oTbl.Rows.Add
                    oTbl.Rows.Add
                    Call FormatRows(oTbl, j)
                End If
                 'Insert the Picture
                ActiveDocument.InlineShapes.AddPicture _
                Filename:=.SelectedItems(i), LinkToFile:=False, _
                SaveWithDocument:=True, Range:=oTbl.Rows(j + 1).Cells(k).Range
                 'Get the Image name for the Caption
                strTxt = Split(.SelectedItems(i), "\")(UBound(Split(.SelectedItems(i), "\")))
                 'Insert the Caption on the row above the picture
                oTbl.Rows(j).Cells(k).Range.Text = Split(strTxt, ".")(0)
            Next
        Else
        End If
    End With
    Application.ScreenUpdating = True
End Function
Function FormatRows(oTbl As Table, x As Long)
    With oTbl
        With .Rows(x + 1)
            .Height = CentimetersToPoints(7)
            .HeightRule = wdRowHeightExactly
            .Range.Style = "Normal"
        End With
        With .Rows(x)
            .Height = CentimetersToPoints(0.75)
            .HeightRule = wdRowHeightExactly
            .Range.Style = "Normal"
        End With
    End With
End Function
Function ReadProp(sPropName As String) As Variant

Dim bCustom As Boolean
Dim sValue As String

   On Error GoTo ErrHandlerReadProp
   'Try the built-in properties first
   'An error will occur if the property doesn't exist
   sValue = ActiveDocument.BuiltInDocumentProperties(sPropName).Value
   ReadProp = sValue
   Exit Function

ContinueCustom:
   bCustom = True

Custom:
   sValue = ActiveDocument.CustomDocumentProperties(sPropName).Value
   ReadProp = sValue
   Exit Function

ErrHandlerReadProp:
   Err.Clear
   'The boolean bCustom has the value False, if this is the first
   'time that the errorhandler is runned
   If Not bCustom Then
     'Continue to see if the property is a custom documentproperty
     Resume ContinueCustom
   Else
     'The property wasn't found, return an empty string
     ReadProp = ""
     Exit Function
   End If

End Function
Function PropsToTable()
' Document Properties
    Dim oRange As Word.Range
    Dim oProp  As DocumentProperty
    Dim sTmp   As String
     
     '\\ value for table header
    sTmp = "Property Name" & vbTab & "Property Value"
     
     '\\ To continue if a document property has no value set
    On Error Resume Next
     
     '\\ Loop document properties and build tab delimeted string
    For Each oProp In ActiveDocument.BuiltInDocumentProperties
        sTmp = sTmp & vbTab & oProp.Name & vbTab & oProp.Value
    Next
     
     '\\ Set reference to start of document (range)
    Set oRange = ActiveDocument.Range(0, 0)
     
     '\\ Insert the string and covert it to a table
    With oRange
        .InsertAfter Text:=sTmp
        .ConvertToTable Separator:=wdSeparateByTabs, _
        NumColumns:=2, _
        AutoFit:=True
    End With
     
     'Clean up
    Set oRange = Nothing
End Function
Function CustomPropsToTable()
' Document Custom Properties
    Dim oRange As Word.Range
    Dim oProp  As DocumentProperty
    Dim sTmp   As String
     
     '\\ value for table header
    sTmp = "Property Name" & vbTab & "Property Value"
     
     '\\ To continue if a document property has no value set
    On Error Resume Next
     
     '\\ Loop document properties and build tab delimeted string
    For Each oProp In ActiveDocument.CustomDocumentProperties
        sTmp = sTmp & vbTab & oProp.Name & vbTab & oProp.Value
    Next
     
     '\\ Set reference to start of document (range)
    Set oRange = ActiveDocument.Range(0, 0)
     
     '\\ Insert the string and covert it to a table
    With oRange
        .InsertAfter Text:=sTmp
        .ConvertToTable Separator:=wdSeparateByTabs, _
        NumColumns:=2, _
        AutoFit:=True
    End With
     
     'Clean up
    Set oRange = Nothing
End Function
Function replace_hdrs()
Dim v As String, lstring As String, strLVL As String
Dim i As Long
Dim myarr As Variant
myarr = Array("H1", "H2", "H3", "Heading Apx 1 Surescripts")
' capture and format the heading
For i = 0 To UBound(myarr)
    Selection.HomeKey wdStory
    On Error Resume Next
    Selection.Find.Style = ActiveDocument.Styles(myarr(i))
    Selection.Find.Text = ""
    Do While Selection.Find.Execute = True

        v = Selection.Range.ListFormat.ListValue
        lstring = Selection.Range.ListFormat.ListString
        strLVL = Selection.Range.ListFormat.ListLevelNumber
        'MsgBox "List value " & v & " is represented by the string " & lstring & " Level " & strLVL
        While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
            Selection.MoveEnd Unit:=wdCharacter, Count:=-1
        Wend
        If strLVL = 1 Then
            Selection.InsertBefore "<br style='clear: both; page-break-before: always;'><p><strong><span style='font-family: Arial; font-size: 18px;'>" & lstring & "  "
        End If
        If strLVL = 2 Then
            Selection.InsertBefore "<p><strong><span style='font-family: Arial; font-size: 16px;'>" & lstring & "  "
        End If
        If strLVL = 3 Then
            Selection.InsertBefore "<p><strong><span style='font-family: Arial; font-size: 13px;'>" & lstring & "  "
        End If
    Selection.InsertAfter "</span></strong></p>"
    Selection.MoveRight wdCharacter, 1
    Loop
Next
' remove the formatting
For i = 0 To UBound(myarr)
    Selection.HomeKey wdStory
    'Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles(myarr(i))
    Selection.Find.Text = ""
    Do While Selection.Find.Execute = True

        v = Selection.Range.ListFormat.ListValue
        lstring = Selection.Range.ListFormat.ListString
        'MsgBox "List value " & v & " is represented by the string " & lstring
        While Selection.Characters.Last = " " Or Selection.Characters.Last = vbCr
            Selection.MoveEnd Unit:=wdCharacter, Count:=-1
        Wend
        Selection.Find.Replacement.ClearFormatting
        Selection.Style = -1
        Selection.MoveRight wdCharacter, 1
    Loop
Next
End Function
Function convertback(strTable)
' convert tables
Dim tTable As Table
Dim oRange As Word.Range
    For Each tTable In ActiveDocument.Tables
        If tTable.ID = strTable Then
            Set oRange = ActiveDocument.Bookmarks(strTable).Range
            oRange.ConvertToTable Separator:=wdSeparateByTabs
        End If
        Set oRange = Nothing
        Exit Function
    Next tTable
End Function
Function replaceSpace(sText As String)
    replaceSpace = Replace(sText, " ", "_")
End Function
Function DeleteUnusedStyles()
    Dim oStyle As Style
    For Each oStyle In ActiveDocument.Styles
        'Only check out non-built-in styles
        If oStyle.BuiltIn = False Then
            With ActiveDocument.Content.Find
                .ClearFormatting
                .Style = oStyle.NameLocal
                .Execute FindText:=””, Format:=True
                If .Found = False Then oStyle.Delete
            End With
        End If
    Next oStyle
End Function
Function GetLevel(strItem As String) As Integer
    ' Return the heading level of a header from the
    ' array returned by Word.

    ' The number of leading spaces indicates the
    ' outline level (2 spaces per level: H1 has
    ' 0 spaces, H2 has 2 spaces, H3 has 4 spaces.

    Dim strTemp As String
    Dim strOriginal As String
    Dim intDiff As Integer

    ' Get rid of all trailing spaces.
    strOriginal = RTrim$(strItem)

    ' Trim leading spaces, and then compare with
    ' the original.
    strTemp = LTrim$(strOriginal)

    ' Functiontract to find the number of
    ' leading spaces in the original string.
    intDiff = Len(strOriginal) - Len(strTemp)
    GetLevel = (intDiff / 2) + 1
End Function
Function CreateTOC()
    'Dim docOutline As Word.Document
    Dim docSource As Word.Document
    Dim rng As Word.Range
    Dim strFootNum() As Integer
    Dim astrHeadings As Variant
    Dim strText As String
    Dim intLevel As Integer
    Dim intItem As Integer
    Dim minLevel As Integer
    Dim tabStops As Variant

    Set docSource = ActiveDocument
    'Set docOutline = Documents.Add
    
    minLevel = 5  'levels above this value won't be copied.

    ' Content returns only the
    ' main body of the document, not
    ' the headers and footer.
    Set rng = docSource.Content
    astrHeadings = docSource.GetCrossReferenceItems(wdRefTypeHeading)

    docSource.Select
    ReDim strFootNum(0 To UBound(astrHeadings))
    For i = 1 To UBound(astrHeadings)
        With Selection.Find
            .Text = Trim(astrHeadings(i))
            .Wrap = wdFindContinue
        End With

        If Selection.Find.Execute = True Then
        On Error Resume Next
            strFootNum(i) = Selection.Information(wdActiveEndPageNumber)
        End If
        Selection.Move
    Next

    docSource.Select

    With Selection.Paragraphs.tabStops
        '.Add Position:=InchesToPoints(2), Alignment:=wdAlignTabLeft
        .Add Position:=InchesToPoints(6), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
    End With

j = 4

    For intItem = LBound(astrHeadings) To UBound(astrHeadings)
        ' Get the text and the level.
        ' strText = Trim$(astrHeadings(intItem))
        intLevel = GetLevel(CStr(astrHeadings(intItem)))
        ' Test which heading is selected and indent accordingly

        If intLevel <= minLevel Then
                If intLevel = "1" Then

                    strText = " " & Trim$(astrHeadings(intItem)) & vbTab & j & vbCr
                End If
                If intLevel = "2" Then
                    strText = "   " & Trim$(astrHeadings(intItem)) & vbTab & j & vbCr
                End If
                If intLevel = "3" Then
                    strText = "      " & Trim$(astrHeadings(intItem)) & vbTab & j & vbCr
                End If
                If intLevel = "4" Then
                    strText = "         " & Trim$(astrHeadings(intItem)) & vbTab & j & vbCr
                End If
                If intLevel = "5" Then
                    strText = "            " & Trim$(astrHeadings(intItem)) & vbTab & j & vbCr
                End If
            ' Add the text to the document.
            rng.InsertAfter strText & vbLf
            docSource.SelectAllEditableRanges
            ' tab stop to set at 15.24 cm
            'With Selection.Paragraphs.tabStops
            '    .Add Position:=InchesToPoints(6), _
            '    Leader:=wdTabLeaderDots, Alignment:=wdAlignTabRight
            '    .Add Position:=InchesToPoints(2), Alignment:=wdAlignTabCenter
            'End With
            rng.Collapse wdCollapseEnd
        End If
    If j <= UBound(astrHeadings) Then j = j + 1
    Next intItem


End Function
Sub tidy(ByVal filesaveas As String)
' Experimental
' Attempting to properly format the output using tidy
'Const TIDY_PROGRAM_FILE = "C:\Users\rob.vance\Documents\Validator\tidy.exe"

Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1

filesaveas = "Regulatory Compliance Policy.html"

'wsh.Run "cmd.exe /S /C " & TIDY_PROGRAM_FILE & " --output-xhtml y --indent 'auto' --indent-spaces '2' --wrap '90' -f " & TIDY_ERROR_FILE & " -m " & filesaveas, windowStyle, waitOnReturn
wsh.Run "cmd.exe /S /C C:\Users\rob.vance\Documents\Validator\tidy.exe --output-xhtml y --indent 'auto' --indent-spaces '2' --wrap '90' -f C:\Users\rob.vance\Documents\Validator\tidy_errors.txt -m " & filesaveas, windowStyle, waitOnReturn

End Sub

'Detect Endnote
Function EndnotesExist() As Boolean

    Dim myStoryRange As Range

    For Each myStoryRange In ActiveDocument.StoryRanges

        If myStoryRange.StoryType = wdEndnotesStory Then

            EndnotesExist = True

            Exit For

        End If

    Next myStoryRange

End Function

'Detect Footnote
Function FootnotesExist() As Boolean

    Dim myStoryRange As Range

    For Each myStoryRange In ActiveDocument.StoryRanges

        If myStoryRange.StoryType = wdFootnotesStory Then

            FootnotesExist = True

            Exit For

        End If

    Next myStoryRange

End Function



