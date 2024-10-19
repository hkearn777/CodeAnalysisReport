Imports System.IO
Imports System.Windows.Forms.LinkLabel
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word

Public Class Form1
  ' Create the Code Analysis Report based on the model spreadsheet.
  ' The output will be MS Word Docs. One document per program name.
  ' The documents will be named: programName_CAR.docx
  '----------------------------------------------------------------------
  ' Change-history.
  Dim ProgramVersion As String = "v0.0"
  ' 2024/10/19 HK v0.0 New Code
  '----------------------------------------------------------------------
  Dim objExcel As New Microsoft.Office.Interop.Excel.Application
  ' Model 
  Dim workbook As Microsoft.Office.Interop.Excel.Workbook
  Dim FilesWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim ProgramsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim SummaryWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim JobsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim JobCommentsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim RecordsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim theWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim FieldsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim CommentsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim EXECSQLWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim EXECCICSWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim IMSWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim DataComWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim ScreenMapWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim CallsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim StatsWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim LibrariesWorksheet As Microsoft.Office.Interop.Excel.Worksheet

  Dim rngSummaryName As Microsoft.Office.Interop.Excel.Range
  Dim rngJobs As Microsoft.Office.Interop.Excel.Range
  Dim rngJobComments As Microsoft.Office.Interop.Excel.Range
  Dim rngPrograms As Microsoft.Office.Interop.Excel.Range
  Dim rngFiles As Microsoft.Office.Interop.Excel.Range
  Dim rngRecordsName As Microsoft.Office.Interop.Excel.Range
  Dim rngFieldsName As Microsoft.Office.Interop.Excel.Range
  Dim rngComments As Microsoft.Office.Interop.Excel.Range
  Dim rngEXECSQL As Microsoft.Office.Interop.Excel.Range
  Dim rngEXECCICS As Microsoft.Office.Interop.Excel.Range
  Dim rngIMS As Microsoft.Office.Interop.Excel.Range
  Dim rngDataCom As Microsoft.Office.Interop.Excel.Range
  Dim rngCalls As Microsoft.Office.Interop.Excel.Range
  Dim rngScreenMap As Microsoft.Office.Interop.Excel.Range
  Dim rngStats As Microsoft.Office.Interop.Excel.Range
  Dim rngLibraries As Microsoft.Office.Interop.Excel.Range

  Dim DefaultFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault
  Dim SetAsReadOnly = Microsoft.Office.Interop.Excel.XlFileAccess.xlReadOnly

  ' Load the Word references
  Dim oWord As New Microsoft.Office.Interop.Word.Application
  Dim oDoc As Word.Document
  Dim oPara1 As Word.Paragraph
  Dim oTable As Word.Table

  Dim ListOfPrograms As New List(Of String)

  Dim Delimiter As String = "|"
  Dim NumOfWordDocBuilt As Integer = 0
  '  Index pointers to the program names on the Tabs
  Dim CommentIndexes As New Dictionary(Of String, String)
  Dim FilesIndexes As New Dictionary(Of String, String)
  Dim RecordsIndexes As New Dictionary(Of String, String)
  Dim FieldsIndexes As New Dictionary(Of String, String)
  Dim TablesIndexes As New Dictionary(Of String, String)


  Private Sub btnCodeAnalysisReport_Click(sender As Object, e As EventArgs) Handles btnCodeAnalysisReport.Click
    Me.Cursor = Cursors.WaitCursor
    lblStatus.Text = "Looking for Programs..."

    objExcel.Visible = False

    ' Identify the PROGRAMS to be selected based on USAGE (not definition)
    ' TAB: ExecSQL
    '   Column: Statement, contains: CST
    '      and
    '   Column: Table, contains: T332
    workbook = objExcel.Workbooks.Open(txtModelName.Text, True, SetAsReadOnly)
    theWorksheet = workbook.Sheets.Item(txtTab.Text)
    theWorksheet.Activate()
    Dim MaxRows As Long = theWorksheet.UsedRange.Rows(theWorksheet.UsedRange.Rows.Count).row
    Dim MaxCols As Long = theWorksheet.UsedRange.Columns.Count
    ' find the Column positions
    Dim col1 As Integer = 0
    Dim col2 As Integer = 0
    For col As Integer = 1 To MaxCols
      Dim colName As String = theWorksheet.Cells(1, col).value2
      If colName = txtColumn1.Text Then
        col1 = col
      End If
      If txtColumn2.TextLength > 0 And colName = txtColumn2.Text Then
        col2 = col
      End If
    Next
    If col1 = 0 Then
      MessageBox.Show("Column1 " & txtColumn1.Text & " not found on tab " & txtTab.Text)
      Exit Sub
    End If
    If txtColumn2.TextLength > 0 And col2 = 0 Then
      MessageBox.Show("Column2 " & txtColumn2.Text & " not found on tab " & txtTab.Text)
      Exit Sub
    End If
    ' find all programs that match the filters
    For Row As Integer = 2 To MaxRows
      ' get the filter values
      Dim col1Value As String = theWorksheet.Cells(Row, col1).value2
      If col1Value Is Nothing Then
        Continue For
      End If
      If col1Value.Length = 0 Then
        Continue For
      End If
      ' match to the contains filter
      If Not col1Value.Contains(txtContains1.Text) Then
        Continue For
      End If
      Dim col2Value As String = ""
      If col2 > 0 Then
        col2Value = theWorksheet.Cells(Row, col2).value2
        If col2Value Is Nothing Then
          Continue For
        End If
        If col2Value.Length = 0 Then
          Continue For
        End If
      End If
      If col2 > 0 And col2Value.Contains(txtContains2.Text) Then
      Else
        Continue For
      End If
      ' this row matches the filter, save it
      Dim programName As String = (theWorksheet.Cells(Row, 2).Value2)
      ' add to list unique
      If ListOfPrograms.IndexOf(programName) = -1 Then
        ListOfPrograms.Add(programName)
      End If
    Next
    lbIndexes.Items.Add("Programs:" & ListOfPrograms.Count)
    '
    ' Load the indexes for select tabs
    '
    lblStatus.Text = "Loading comments indexes..."
    Dim resultCnt As Integer = GetCommentProgramIndexes()
    lbIndexes.Items.Add("Comments:" & CommentIndexes.Count)
    '
    lblStatus.Text = "Loading Files indexes..."
    resultCnt = GetFilesProgramIndexes()
    lbIndexes.Items.Add("Files:" & FilesIndexes.Count)
    '
    lblStatus.Text = "Loading Records indexes..."
    resultCnt = GetRecordsProgramIndexes()
    lbIndexes.Items.Add("Records:" & RecordsIndexes.Count)
    '
    lblStatus.Text = "Loading Fields indexes..."
    resultCnt = GetFieldsProgramIndexes()
    lbIndexes.Items.Add("Fields:" & FieldsIndexes.Count)
    '
    lblStatus.Text = "Loading Tables indexes..."
    resultCnt = GetTablesProgramIndexes()
    lbIndexes.Items.Add("Tables:" & TablesIndexes.Count)
    '
    '
    ' Create all the WORD documents
    '
    lblStatus.Text = "Creating" & ListOfPrograms.Count & " Word Documents..."
    For Each programName In ListOfPrograms
      Call CreateWordDocument(programName)
      NumOfWordDocBuilt += 1
      lbIndexes.Items.Add(NumOfWordDocBuilt & " " & programName & " Built")
      lblStatus2.Text = NumOfWordDocBuilt & " " & programName & " Built"
      'Exit For
    Next
    lblStatus.Text = "Complete. Number of Word Docs built:" & NumOfWordDocBuilt


    workbook.Close()
    objExcel.Quit()
    oWord.Quit()
    Me.Cursor = Cursors.Default
    MessageBox.Show("Complete")

  End Sub

  Sub CreateWordDocument(ByRef thePgm As String)
    ' create a CAR (code analysis report)
    oWord.Visible = False
    oDoc = oWord.Documents.Add

    'Dim aHLink = oDoc.Hyperlinks.Add(Anchor:=Selection.Range, Address:="https://forms")

    For Each section As Word.Section In oWord.ActiveDocument.Sections
      Dim headerRange As Word.Range = section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range
      'headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage)
      headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
      headerRange.Text = "Lowe's Corporation Code Analysis Report"
    Next


    'Insert a paragraph at the beginning of the document.
    oPara1 = oDoc.Content.Paragraphs.Add
    oPara1.Range.Text = "Analysis for: " & txtColumn1.Text & " or " & txtContains1.Text
    oPara1.Range.Font.Bold = True
    oPara1.Format.SpaceAfter = 24    '24 pt spacing after paragraph.
    oPara1.Range.InsertParagraphAfter()

    'State the Program Name.
    oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    oPara1.Range.Font.Bold = True
    oPara1.Range.Text = "Program: " & thePgm
    oPara1.Format.SpaceAfter = 6
    oPara1.Range.InsertParagraphAfter()

    ' Comments / purpose.
    oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    oPara1.Range.Font.Bold = False
    oPara1.Range.Text = "Comments: " & GrabComments(thePgm)
    oPara1.Format.SpaceAfter = 4
    oPara1.Range.InsertParagraphAfter()

    ' Files / record / copybook
    oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    oPara1.Range.Font.Bold = True
    oPara1.Range.Text = "Files / Records / Copybooks Overview"
    oPara1.Format.SpaceAfter = 4
    oPara1.Range.InsertParagraphAfter()

    'Insert a 3 x n table, fill it with data, and make the first row bold
    Dim ListOfFiles As List(Of String) = GrabFileRecordCopybook(thePgm)
    Dim r As Integer = 1, c As Integer = 1
    oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, ListOfFiles.Count + 1, 3)
    oTable.AllowAutoFit = True
    oTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)
    oTable.Range.ParagraphFormat.SpaceAfter = 6
    oTable.Rows.Item(1).Range.Font.Bold = True
    oTable.Columns.AutoFit()
    oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle
    oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle
    oTable.Cell(1, 1).Range.Text = "File/Table"
    oTable.Cell(1, 2).Range.Text = "Record Name"
    oTable.Cell(1, 3).Range.Text = "Copybook"
    For Each file In ListOfFiles
      Dim fileEntry As String() = file.Split(Delimiter)
      r += 1
      oTable.Rows.Item(r).Range.Font.Bold = False
      oTable.Cell(r, 1).Range.Text = fileEntry(0)     'File/Table
      oTable.Cell(r, 2).Range.Text = fileEntry(1)     'Record Name
      oTable.Cell(r, 3).Range.Text = fileEntry(2)     'Copybook
    Next
    '
    ' Fields layouts
    '
    oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    oPara1.Range.Font.Bold = True
    oPara1.Range.Text = "Field Layout details:" & vbCrLf
    oPara1.Format.SpaceAfter = 4
    oPara1.Range.InsertParagraphAfter()

    Dim ListOfFields As List(Of String) = GrabFields(thePgm)
    Dim RecordNameCounts As Dictionary(Of String, Integer) = CountRecordNames(ListOfFields)
    For Each entry In ListOfFiles
      Dim fileEntry As String() = entry.Split(Delimiter)
      oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
      oPara1.Range.Font.Bold = True
      oPara1.Range.Text = fileEntry(0) & " / " & fileEntry(1) & " / " & fileEntry(2) & ":" & vbCrLf
      oPara1.Format.SpaceAfter = 4
      oPara1.Range.InsertParagraphAfter()
      'Insert a 7 x n table, fill it with data, and make the first row bold
      r = 1
      c = 1
      Dim CountOfFields As Integer = 0
      Dim result As String = ""
      Dim searchName As String = fileEntry(0) & Delimiter & fileEntry(1)
      If RecordNameCounts.TryGetValue(searchName, CountOfFields) Then
      Else
        MessageBox.Show("Error getting recordNameCounts for:" & searchName)
      End If
      oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, CountOfFields + 1, 7)
      'Dim ts As TableLayoutStyleCollection
      oTable.Range.ParagraphFormat.SpaceAfter = 6
      oTable.Rows.Item(1).Range.Font.Bold = True
      oTable.AllowAutoFit = True
      oTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)
      oTable.Columns.AutoFit()
      oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle
      oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle
      oTable.Cell(1, 1).Range.Text = "Seq"
      oTable.Cell(1, 2).Range.Text = "Lvl"
      oTable.Cell(1, 3).Range.Text = "Field Name"
      'oTable.Cell(1, 3).Column.AutoFit()
      oTable.Cell(1, 4).Range.Text = "Picture"
      oTable.Cell(1, 5).Range.Text = "Start"
      oTable.Cell(1, 6).Range.Text = "End"
      oTable.Cell(1, 7).Range.Text = "Length"
      oTable.Cell(1, 1).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
      oTable.Cell(1, 2).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
      oTable.Cell(1, 3).SetWidth(200, WdRulerStyle.wdAdjustSameWidth)
      oTable.Cell(1, 4).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
      oTable.Cell(1, 5).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
      oTable.Cell(1, 6).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
      oTable.Cell(1, 7).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
      For Each fieldEntries In ListOfFields
        Dim fieldEntry As String() = fieldEntries.Split(Delimiter)
        Dim fileKey As String = fileEntry(0) & Delimiter & fileEntry(1)
        Dim fieldKey As String = fieldEntry(0) & Delimiter & fieldEntry(1)
        If fieldKey <> fileKey Then         'check file / recordname
          Continue For
        End If
        r += 1
        If r > CountOfFields + 1 Then
          MessageBox.Show("r > count:file/record:" & fieldKey)
        End If
        oTable.Rows.Item(r).Range.Font.Bold = False
        oTable.Cell(r, 1).Range.Text = fieldEntry(3)     'seq
        oTable.Cell(r, 2).Range.Text = fieldEntry(4)     'level
        oTable.Cell(r, 3).Range.Text = fieldEntry(5)     'field name
        oTable.Cell(r, 4).Range.Text = fieldEntry(6)     'picture
        oTable.Cell(r, 5).Range.Text = fieldEntry(7)     'start
        oTable.Cell(r, 6).Range.Text = fieldEntry(8)     'end
        oTable.Cell(r, 7).Range.Text = fieldEntry(9)     'length
        oTable.Cell(r, 1).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
        oTable.Cell(r, 2).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
        oTable.Cell(r, 3).SetWidth(200, WdRulerStyle.wdAdjustSameWidth)
        oTable.Cell(r, 4).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
        oTable.Cell(r, 5).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
        oTable.Cell(r, 6).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)
        oTable.Cell(r, 7).SetWidth(50, WdRulerStyle.wdAdjustSameWidth)

      Next

    Next

    '
    ' grab DB tables used
    '
    oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    oPara1.Range.Font.Bold = True
    oPara1.Range.Text = vbCrLf & "Tables"
    oPara1.Format.SpaceAfter = 4
    oPara1.Range.InsertParagraphAfter()
    '
    'Insert a 2 x n table, fill it with data, and make the first row bold
    Dim ListOfTables As List(Of String) = GrabTables(thePgm)
    ListOfTables.Sort()
    r = 1
    c = 1
    oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, ListOfTables.Count + 1, 2)
    oTable.AllowAutoFit = True
    oTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)
    oTable.Range.ParagraphFormat.SpaceAfter = 6
    oTable.Rows.Item(1).Range.Font.Bold = True
    oTable.Columns.AutoFit()
    oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle
    oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle
    oTable.Cell(1, 1).Range.Text = "Table Name"
    oTable.Cell(1, 2).Range.Text = "SQL Command"
    For Each table In ListOfTables
      Dim fileEntry As String() = table.Split(Delimiter)
      r += 1
      oTable.Rows.Item(r).Range.Font.Bold = False
      oTable.Cell(r, 1).Range.Text = fileEntry(0)     'Table
      oTable.Cell(r, 2).Range.Text = fileEntry(1)     'execSQL
    Next

    '
    ' grab cursor text
    '
    oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    oPara1.Range.Font.Bold = True
    oPara1.Range.Text = vbCrLf & "Table Cursors"
    oPara1.Format.SpaceAfter = 4
    oPara1.Range.InsertParagraphAfter()
    '
    'Insert a 3 x n table, fill it with data, and make the first row bold
    Dim ListOfCursors As List(Of String) = GrabCursors(thePgm)
    'ListOfCursors.Sort()
    r = 1
    c = 1
    oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, ListOfCursors.Count + 1, 3)
    oTable.AllowAutoFit = True
    oTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)
    oTable.Range.ParagraphFormat.SpaceAfter = 6
    oTable.Rows.Item(1).Range.Font.Bold = True
    oTable.Columns.AutoFit()
    oTable.Cell(1, 1).Range.Text = "Cursor Name"
    oTable.Cell(1, 2).Range.Text = "SQL Command"
    oTable.Cell(1, 3).Range.Text = "Statement"
    oTable.Cell(1, 1).SetWidth(100, WdRulerStyle.wdAdjustSameWidth)
    oTable.Cell(1, 2).SetWidth(75, WdRulerStyle.wdAdjustSameWidth)
    oTable.Cell(1, 3).SetWidth(300, WdRulerStyle.wdAdjustSameWidth)

    oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle
    oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle

    For Each table In ListOfCursors
      Dim fileEntry As String() = table.Split(Delimiter)
      r += 1
      oTable.Rows.Item(r).Range.Font.Bold = False
      oTable.Cell(r, 1).Range.Text = fileEntry(0)     'Table
      oTable.Cell(r, 2).Range.Text = fileEntry(1)     'execSQL
      oTable.Cell(r, 3).Range.Text = fileEntry(2)     'Statement
      oTable.Cell(r, 1).SetWidth(100, WdRulerStyle.wdAdjustSameWidth)
      oTable.Cell(r, 2).SetWidth(75, WdRulerStyle.wdAdjustSameWidth)
      oTable.Cell(r, 3).SetWidth(300, WdRulerStyle.wdAdjustSameWidth)
    Next


    ' grab business rules
    oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    oPara1.Range.Font.Bold = True
    oPara1.Range.Text = vbCrLf & "Business Rules"
    oPara1.Format.SpaceAfter = 4
    oPara1.Range.InsertParagraphAfter()

    ' grab flowcharts
    oPara1 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    oPara1.Range.Font.Bold = True
    oPara1.Range.Text = vbCrLf & "Flowcharts"
    oPara1.Format.SpaceAfter = 4
    oPara1.Range.InsertParagraphAfter()

    'save the file
    Dim ProgramsFileName As String = txtWorkFolder.Text & "\" & thePgm & "_CAR.docx"
    oDoc.SaveAs2(ProgramsFileName)
    oDoc.Close()

  End Sub
  Function CountRecordNames(ByRef ListOfFields As List(Of String)) As Dictionary(Of String, Integer)
    ' count the number of fields for each record-name
    ' presume recordnames are grouped together
    Dim recordDict As New Dictionary(Of String, Integer)
    Dim RecordFound As Boolean = False
    Dim lastRecordName As String = ""
    Dim cnt As Integer = 1
    Dim firstRow As String() = ListOfFields(0).Split(Delimiter)
    lastRecordName = firstRow(0) & Delimiter & firstRow(1)

    For x As Integer = 1 To ListOfFields.Count - 1
      Dim fieldEntry As String() = ListOfFields(x).Split(Delimiter)
      Dim currRecord As String = fieldEntry(0) & Delimiter & fieldEntry(1)
      If currRecord = lastRecordName Then
        cnt += 1
        Continue For
      End If
      recordDict.Add(lastRecordName, cnt)
      lastRecordName = currRecord
      cnt = 1
    Next
    If cnt > 0 Then
      recordDict.Add(lastRecordName, cnt)
    End If
    Return recordDict
  End Function

  Function GetCommentProgramIndexes() As Integer
    CommentsWorksheet = workbook.Sheets.Item("Comments")
    CommentsWorksheet.Activate()

    Dim MaxRows As Long = CommentsWorksheet.UsedRange.Rows(CommentsWorksheet.UsedRange.Rows.Count).row
    For Row As Integer = 2 To MaxRows
      Dim programName As String = CommentsWorksheet.Cells.Range("B" & Row).Value2
      Dim theKey As String = programName
      Dim theValue As String = LTrim(Str(Row))
      If Not CommentIndexes.ContainsKey(theKey) Then
        CommentIndexes.Add(theKey, theValue)
      End If
    Next
    Return CommentIndexes.Count
  End Function
  Function GetFilesProgramIndexes() As Integer
    FilesWorksheet = workbook.Sheets.Item("Files")
    FilesWorksheet.Activate()
    Dim MaxRows As Long = FilesWorksheet.UsedRange.Rows(FilesWorksheet.UsedRange.Rows.Count).row
    For Row As Integer = 2 To MaxRows
      Dim programName As String = FilesWorksheet.Cells.Range("F" & Row).Value2
      If programName Is Nothing Then
        Continue For
      End If
      If programName.Length = 0 Then
        Continue For
      End If
      Dim theKey As String = programName
      Dim theValue As String = LTrim(Str(Row))
      If Not FilesIndexes.ContainsKey(theKey) Then
        FilesIndexes.Add(theKey, theValue)
      End If
    Next
    Return FilesIndexes.Count
  End Function
  Function GetRecordsProgramIndexes() As Integer
    RecordsWorksheet = workbook.Sheets.Item("Records")
    RecordsWorksheet.Activate()
    Dim MaxRows As Long = RecordsWorksheet.UsedRange.Rows(RecordsWorksheet.UsedRange.Rows.Count).row
    For Row As Integer = 2 To MaxRows
      Dim programName As String = RecordsWorksheet.Cells.Range("B" & Row).Value2
      If programName Is Nothing Then
        Continue For
      End If
      If programName.Length = 0 Then
        Continue For
      End If
      Dim theKey As String = programName
      Dim theValue As String = LTrim(Str(Row))
      If Not RecordsIndexes.ContainsKey(theKey) Then
        RecordsIndexes.Add(theKey, theValue)
      End If
    Next
    Return RecordsIndexes.Count
  End Function
  Function GetFieldsProgramIndexes() As Integer
    FieldsWorksheet = workbook.Sheets.Item("Fields")
    FieldsWorksheet.Activate()
    Dim MaxRows As Long = FieldsWorksheet.UsedRange.Rows(FieldsWorksheet.UsedRange.Rows.Count).row
    For Row As Integer = 2 To MaxRows
      Dim programName As String = FieldsWorksheet.Cells.Range("B" & Row).Value2
      If programName Is Nothing Then
        Continue For
      End If
      If programName.Length = 0 Then
        Continue For
      End If
      Dim theKey As String = programName
      Dim theValue As String = LTrim(Str(Row))
      If Not FieldsIndexes.ContainsKey(theKey) Then
        FieldsIndexes.Add(theKey, theValue)
      End If
    Next
    Return FieldsIndexes.Count
  End Function
  Function GetTablesProgramIndexes() As Integer
    EXECSQLWorksheet = workbook.Sheets.Item("ExecSQL")
    EXECSQLWorksheet.Activate()
    Dim MaxRows As Long = EXECSQLWorksheet.UsedRange.Rows(EXECSQLWorksheet.UsedRange.Rows.Count).row
    For Row As Integer = 2 To MaxRows
      Dim programName As String = EXECSQLWorksheet.Cells.Range("B" & Row).Value2
      If programName Is Nothing Then
        Continue For
      End If
      If programName.Length = 0 Then
        Continue For
      End If
      Dim theKey As String = programName
      Dim theValue As String = LTrim(Str(Row))
      If Not TablesIndexes.ContainsKey(theKey) Then
        TablesIndexes.Add(theKey, theValue)
      End If
    Next
    Return TablesIndexes.Count
  End Function
  Function GrabComments(ByRef thePgm As String) As String
    ' get / format coments as one string
    CommentsWorksheet = workbook.Sheets.Item("Comments")
    CommentsWorksheet.Activate()
    ' Get starting index
    Dim result As String = ""
    If CommentIndexes.TryGetValue(thePgm, result) Then
    Else
      MessageBox.Show("wow! No comments for program:" & thePgm & ": not found!")
      Return "No comments found for program"
    End If
    Dim startIndex As Integer = Val(result)

    Dim WholeComments As String = ""
    Dim MaxRows As Long = CommentsWorksheet.UsedRange.Rows(CommentsWorksheet.UsedRange.Rows.Count).row
    Dim FoundStarted As Boolean = False
    For recordsRow As Integer = startIndex To MaxRows
      Dim Row As String = LTrim(Str(recordsRow))
      Dim programName As String = CommentsWorksheet.Cells.Range("B" & Row).Value2
      Dim divisionName As String = CommentsWorksheet.Cells.Range("D" & Row).Value2
      Dim lineNumber As String = CommentsWorksheet.Cells.Range("E" & Row).Value2
      Dim comment As String = CommentsWorksheet.Cells.Range("F" & Row).Value2
      If Not programName.Equals(thePgm) Then
        If FoundStarted Then        'so I don't have to search the whole sheet again and again
          Exit For
        End If
        Continue For
      End If
      FoundStarted = True
      If Val(lineNumber) <= 20 Then
        WholeComments &= comment.Replace(vbCrLf, " ") & vbCrLf
      Else
        Exit For
      End If
    Next
    Return WholeComments.Replace(Chr(34), "")
  End Function
  Function GrabFileRecordCopybook(ByRef thePgm As String) As List(Of String)
    ' get Names of files, record names, and copybooks
    Dim ListOfFiles As New List(Of String)
    RecordsWorksheet = workbook.Sheets.Item("Records")
    RecordsWorksheet.Activate()
    ' Get starting index
    Dim result As String = ""
    If RecordsIndexes.TryGetValue(thePgm, result) Then
    Else
      MessageBox.Show("wow! No File/Record/Copybook for program:" & thePgm & ":!")
      Return Nothing
    End If
    Dim startIndex As Integer = Val(result)

    Dim WholeRecords As String = ""
    Dim MaxRows As Long = RecordsWorksheet.UsedRange.Rows(RecordsWorksheet.UsedRange.Rows.Count).row
    Dim FoundStarted As Boolean = False
    For Row As Integer = startIndex To MaxRows
      Dim programName As String = RecordsWorksheet.Cells(Row, 2).Value2
      Dim fileName As String = RecordsWorksheet.Cells(Row, 3).Value2
      Dim recordName As String = RecordsWorksheet.Cells(Row, 6).Value2
      Dim copybook As String = RecordsWorksheet.Cells(Row, 7).Value2
      If Not programName.Equals(thePgm) Then
        If FoundStarted Then        'so I don't have to search the whole sheet again and again
          Exit For
        End If
        Continue For
      End If
      FoundStarted = True
      ListOfFiles.Add(fileName & Delimiter & recordName & Delimiter & copybook)
    Next
    Return ListOfFiles
  End Function
  Function GrabFields(ByRef thePgm As String) As List(Of String)
    ' get Field names and details for the given pgm name
    Dim ListOfFields As New List(Of String)
    FieldsWorksheet = workbook.Sheets.Item("Fields")
    FieldsWorksheet.Activate()
    ' Get starting index
    Dim result As String = ""
    If FieldsIndexes.TryGetValue(thePgm, result) Then
    Else
      MessageBox.Show("wow! No Fields for program:" & thePgm & ":!")
      Return Nothing
    End If
    Dim startIndex As Integer = Val(result)

    Dim WholeFields As String = ""
    Dim MaxRows As Long = FieldsWorksheet.UsedRange.Rows(FieldsWorksheet.UsedRange.Rows.Count).row
    Dim FoundStarted As Boolean = False

    For Row As Integer = startIndex To MaxRows
      Dim programName As String = FieldsWorksheet.Cells(Row, 2).Value2
      Dim fileName As String = FieldsWorksheet.Cells(Row, 3).Value2
      Dim recordName As String = FieldsWorksheet.Cells(Row, 6).Value2
      Dim copybook As String = FieldsWorksheet.Cells(Row, 7).Value2
      Dim fieldSeq As String = FieldsWorksheet.Cells(Row, 8).Value2
      Dim fieldLevel As String = FieldsWorksheet.Cells(Row, 9).Value2
      Dim fieldName As String = FieldsWorksheet.Cells(Row, 10).Value2
      Dim fieldPicture As String = FieldsWorksheet.Cells(Row, 11).Value2
      Dim fieldStart As String = FieldsWorksheet.Cells(Row, 12).Value2
      Dim fieldEnd As String = FieldsWorksheet.Cells(Row, 13).Value2
      Dim fieldLength As String = FieldsWorksheet.Cells(Row, 14).Value2
      ' only want row that matches program name and record name
      If programName = thePgm Then
        FoundStarted = True
        ListOfFields.Add(fileName & Delimiter &
                           recordName & Delimiter &
                           copybook & Delimiter &
                           fieldSeq & Delimiter &
                           fieldLevel & Delimiter &
                           fieldName & Delimiter &
                           fieldPicture & Delimiter &
                           fieldStart & Delimiter &
                           fieldEnd & Delimiter &
                           fieldLength)
        Continue For
      End If
      ' key does not match
      If FoundStarted Then        'so I don't have to search the whole sheet again and again
        Exit For
      End If
    Next

    Return ListOfFields
  End Function
  Function GrabTables(ByRef thePgm As String) As List(Of String)
    ' get Table names
    Dim ListOfTables As New List(Of String)
    EXECSQLWorksheet = workbook.Sheets.Item("ExecSQL")
    EXECSQLWorksheet.Activate()
    ' Get starting index
    Dim result As String = ""
    If TablesIndexes.TryGetValue(thePgm, result) Then
    Else
      MessageBox.Show("wow! No ExecSQL for program:" & thePgm & ":!")
      Return Nothing
    End If
    Dim startIndex As Integer = Val(result)

    Dim WholeRecords As String = ""
    Dim MaxRows As Long = EXECSQLWorksheet.UsedRange.Rows(EXECSQLWorksheet.UsedRange.Rows.Count).row
    Dim FoundStarted As Boolean = False
    For Row As Integer = startIndex To MaxRows
      Dim programName As String = EXECSQLWorksheet.Cells(Row, 2).Value2
      Dim execSQL As String = EXECSQLWorksheet.Cells(Row, 3).Value2
      Dim tableName As String = EXECSQLWorksheet.Cells(Row, 5).Value2
      If tableName Is Nothing Then
        tableName = ""
      End If
      Dim cursorName As String = EXECSQLWorksheet.Cells(Row, 6).Value2
      If cursorName Is Nothing Then
        cursorName = ""
      End If

      'Dim statement As String = EXECSQLWorksheet.Cells(Row, 7).Value2
      Dim tables As String() = tableName.Split(vbCrLf)
      If Not programName.Equals(thePgm) Then
        If FoundStarted Then        'so I don't have to search the whole sheet again and again
          Exit For
        End If
        Continue For
      End If
      FoundStarted = True
      For Each tabname In tables
        If ListOfTables.IndexOf(tabname & Delimiter & execSQL) = -1 Then
          ListOfTables.Add(tableName & Delimiter & execSQL)
        End If
      Next
    Next
    Return ListOfTables
  End Function

  Function GrabCursors(ByRef thePgm As String) As List(Of String)
    ' get cursor names
    Dim ListOfCursors As New List(Of String)
    EXECSQLWorksheet = workbook.Sheets.Item("ExecSQL")
    EXECSQLWorksheet.Activate()
    ' Get starting index
    Dim result As String = ""
    If TablesIndexes.TryGetValue(thePgm, result) Then
    Else
      MessageBox.Show("wow! No ExecSQL for program:" & thePgm & ":!")
      Return Nothing
    End If
    Dim startIndex As Integer = Val(result)

    Dim WholeRecords As String = ""
    Dim MaxRows As Long = EXECSQLWorksheet.UsedRange.Rows(EXECSQLWorksheet.UsedRange.Rows.Count).row
    Dim FoundStarted As Boolean = False
    For Row As Integer = startIndex To MaxRows
      Dim cursorName As String = EXECSQLWorksheet.Cells(Row, 6).Value2
      If cursorName Is Nothing Then
        Continue For
      End If
      Dim programName As String = EXECSQLWorksheet.Cells(Row, 2).Value2
      Dim execSQL As String = EXECSQLWorksheet.Cells(Row, 3).Value2
      Dim statement As String = EXECSQLWorksheet.Cells(Row, 7).Value2
      If cursorName Is Nothing Then
        cursorName = ""
      End If
      If Not programName.Equals(thePgm) Then
        If FoundStarted Then        'so I don't have to search the whole sheet again and again
          Exit For
        End If
        Continue For
      End If
      FoundStarted = True
      ListOfCursors.Add(cursorName & Delimiter & execSQL & Delimiter & statement)
    Next
    Return ListOfCursors
  End Function

  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    lblStatus.Text = "Click 'Create Lowes Analysis' button to proceed"
    lblStatus2.Text = ""
    Me.Text = "Code Analysis Report " & ProgramVersion

  End Sub
End Class
