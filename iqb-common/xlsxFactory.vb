Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml

Public Enum CellTypes
    int
    str
    datetime
    hyperlink
    dec
    text
    formula
End Enum

Public Class ExcelStyleDefs
    Inherits Dictionary(Of xlsxFactory.CellFormatting, UInt32Value)
End Class


Public Class xlsxFactory
    Public Enum CellFormatting
        Null
        ColHeader
        RowHeader1
        RowHeader2
        NumInt
        Num2Dec
        Hyperlink
        MultiLineText
    End Enum

    Private Shared StandardExcelStyleDefs As New ExcelStyleDefs From {
                                                            {CellFormatting.Null, 0UI},
                                                            {CellFormatting.ColHeader, 1UI},
                                                            {CellFormatting.RowHeader1, 2UI},
                                                            {CellFormatting.RowHeader2, 3UI},
                                                            {CellFormatting.NumInt, 4UI},
                                                            {CellFormatting.Num2Dec, 5UI},
                                                            {CellFormatting.Hyperlink, 6UI},
                                                            {CellFormatting.MultiLineText, 7UI}}


    ''' <summary>
    ''' liefert nächst höhere Spalte
    ''' </summary>
    ''' <remarks>nur bis zweistellig</remarks>
    Public Shared Function GetNextColumn(ByVal s As String) As String
        If String.IsNullOrEmpty(s) Then
            s = ""
        Else
            s = s.ToUpper
            Select Case s.Length
                Case 1
                    s = Chr(Asc(s) + 1)
                    If s > "Z" Then s = "AA"
                Case 2
                    s = s(0) + Chr(Asc(s(1)) + 1)
                    If s(1) > "Z" Then s = Chr(Asc(s(0)) + 1) + "A"
                Case Else
                    s = "AAA"
            End Select
        End If

        Return s
    End Function

    Public Shared Function GetPrevColumn(ByVal s As String) As String
        If String.IsNullOrEmpty(s) OrElse s.ToUpper = "A" Then
            s = ""
        Else
            s = s.ToUpper
            Select Case s.Length
                Case 1
                    s = Chr(Asc(s) - 1)
                Case 2
                    s = s(0) + Chr(Asc(s(1)) - 1)
                    If s(1) < "A" Then
                        If s(0) = "A" Then
                            s = "Z"
                        Else
                            s = Chr(Asc(s(0)) - 1) + "Z"
                        End If
                    End If
                Case Else
                    s = "ZZ"
            End Select
        End If

        Return s
    End Function

    '########################################################################################
    '#### http://msdn.microsoft.com/en-us/library/cc861607.aspx#Y1680
    '########################################################################################

    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Public Shared Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function

    ' Given a WorkbookPart, inserts a new worksheet.
    Public Shared Function InsertWorksheet(ByVal localWorkbookPart As WorkbookPart, Optional SheetName As String = "") As WorksheetPart
        ' Add a new worksheet part to the workbook.
        Dim newWorksheetPart As WorksheetPart = localWorkbookPart.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData)
        newWorksheetPart.Worksheet.Save()
        Dim sheets As Sheets = localWorkbookPart.Workbook.GetFirstChild(Of Sheets)()
        Dim relationshipId As String = localWorkbookPart.GetIdOfPart(newWorksheetPart)

        ' Get a unique ID for the new sheet.
        Dim sheetId As UInteger = 1
        If (sheets.Elements(Of Sheet).Count() > 0) Then sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max() + 1

        If String.IsNullOrEmpty(SheetName) Then SheetName = "Tabelle" + sheetId.ToString()

        ' Add the new worksheet and associate it with the workbook.
        Dim sheet As Sheet = New Sheet
        sheet.Id = relationshipId
        sheet.SheetId = sheetId
        sheet.Name = SheetName
        sheets.Append(sheet)
        localWorkbookPart.Workbook.Save()

        Return newWorksheetPart
    End Function


    'cell3.DataType = CellValues.Date;
    'cell3.CellValue = new CellValue(DateTime.Now.ToOADate().ToString());

    Public Shared Sub SetCellValueString(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart,
                                            CellContent As String, Optional CellFormat As CellFormatting = CellFormatting.Null, Optional StyleDef As ExcelStyleDefs = Nothing)
        If StyleDef Is Nothing Then StyleDef = xlsxFactory.StandardExcelStyleDefs
        Dim myCell As Cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart, StyleDef(CellFormat))
        If myCell IsNot Nothing Then
            If String.IsNullOrEmpty(CellContent) Then
                If myCell.CellValue IsNot Nothing Then myCell.CellValue.Remove()
            Else
                myCell.CellValue = New CellValue(CellContent)
                myCell.DataType = New EnumValue(Of CellValues)(CellValues.String)
            End If
        End If
    End Sub

    Public Shared Sub SetCellValueNumeric(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart,
                                            CellContent As Double, Optional Decimals As Boolean = False, Optional StyleDef As ExcelStyleDefs = Nothing)
        If StyleDef Is Nothing Then StyleDef = xlsxFactory.StandardExcelStyleDefs
        Dim myCell As Cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart, IIf(Decimals, StyleDef(CellFormatting.Num2Dec), StyleDef(CellFormatting.NumInt)))
        If myCell IsNot Nothing Then
            myCell.CellValue = New CellValue(CellContent.ToString(System.Globalization.CultureInfo.InvariantCulture))
            myCell.DataType = New EnumValue(Of CellValues)(CellValues.Number)
        End If
    End Sub

    Private Shared Function ExistingColumnIsGreater(ExistingCell As Cell, column As String) As Boolean
        Dim mc As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ExistingCell.CellReference.Value.ToUpper, "^[A-Z]+")
        Dim existingC As String = mc.Item(0).Value
        If existingC.Length = column.Length Then
            Return (String.Compare(existingC, column.ToUpper, True) > 0)
        Else
            Return existingC.Length > column.Length
        End If
    End Function

    Public Shared Function GetMaxRowIndex(ByVal worksheetPart As WorksheetPart) As Integer
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()

        Return sheetData.Elements(Of Row).Count
    End Function

    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, return it. 
    Private Shared Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger,
                                                    ByVal worksheetPart As WorksheetPart, CellStyle As UInt32Value) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = (columnName + rowIndex.ToString().Trim)

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row
        Dim rowQ = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex)
        If (rowQ.Count() <> 0) Then
            row = rowQ.First()
        Else
            Dim refRow As Row = Nothing
            For Each r As Row In sheetData.Elements(Of Row)()
                If r.RowIndex.Value > rowIndex Then
                    refRow = r
                    Exit For
                End If
            Next
            row = New Row()
            row.RowIndex = rowIndex
            If refRow Is Nothing Then
                sheetData.Append(row)
            Else
                sheetData.InsertBefore(row, refRow)
            End If
        End If

        Dim myCell As Cell = Nothing

        ' If there is not a cell with the specified column name, insert one.  
        Dim colQ = row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference)
        If colQ.Count() > 0 Then
            myCell = colQ.First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If ExistingColumnIsGreater(cell, columnName) Then
                    refCell = cell
                    Exit For
                End If
            Next

            myCell = New Cell With {.CellReference = cellReference}
            myCell.StyleIndex = CellStyle

            row.InsertBefore(myCell, refCell)
        End If
        worksheet.Save()

        Return myCell
    End Function

    'Zur Einfachheit halber und wg Ab
    Public Shared Sub GenerateStyleSheet(WB As WorkbookPart)
        If WB.WorkbookStylesPart IsNot Nothing AndAlso WB.WorkbookStylesPart.Stylesheet IsNot Nothing Then _
            Throw New NotImplementedException("GenerateStyleSheet kann nicht auf ein Workbookpart angewendet werden, der schon Styles enthält. Verwenden Sie 'AddIQBStandardStyles'")

        Dim st As ExcelStyleDefs = AddIQBStandardStyles(WB)
    End Sub

    Public Shared Function AddIQBStandardStyles(WB As WorkbookPart) As ExcelStyleDefs
        If WB.WorkbookStylesPart Is Nothing Then WB.AddNewPart(Of WorkbookStylesPart)()
        If WB.WorkbookStylesPart.Stylesheet Is Nothing Then
            WB.WorkbookStylesPart.Stylesheet = New Stylesheet()
            WB.WorkbookStylesPart.Stylesheet.AddNamespaceDeclaration("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        End If
        Dim mySSh As Stylesheet = WB.WorkbookStylesPart.Stylesheet

        'Nummernformate
        Dim NumberingFormatIDs As New List(Of Integer)
        Dim myNF As NumberingFormats = mySSh.GetFirstChild(Of NumberingFormats)()
        If myNF Is Nothing Then
            mySSh.InsertAt(New NumberingFormats(), 0)
            myNF = mySSh.GetFirstChild(Of NumberingFormats)()
        End If
        NumberingFormatIDs.Add(myNF.ChildElements.Count)
        myNF.Append(New NumberingFormat() With {.FormatCode = "#.##0", .NumberFormatId = 164UI})
        NumberingFormatIDs.Add(myNF.ChildElements.Count)
        myNF.Append(New NumberingFormat() With {.FormatCode = "#,##0.00", .NumberFormatId = 165UI})

        'Fonts
        Dim FontIDs As New List(Of Integer)
        Dim myF As Fonts = mySSh.GetFirstChild(Of Fonts)()
        If myF Is Nothing Then
            mySSh.InsertAfter(New Fonts(), myNF)
            myF = mySSh.GetFirstChild(Of Fonts)()
        End If
        FontIDs.Add(myF.ChildElements.Count)
        myF.Append(New Font(
                        New FontSize() With {.Val = 11},
                        New Color() With {.Rgb = New HexBinaryValue() With {.Value = "0000"}},
                        New FontName() With {.Val = "Calibri"}))
        FontIDs.Add(myF.ChildElements.Count)
        myF.Append(New Font(
                        New Underline(),
                        New FontSize() With {.Val = 11},
                        New Color() With {.Rgb = New HexBinaryValue() With {.Value = "0000"}},
                        New FontName() With {.Val = "Calibri"}))
        FontIDs.Add(myF.ChildElements.Count)
        myF.Append(New Font(
                        New FontSize() With {.Val = 12},
                        New Color() With {.Rgb = New HexBinaryValue() With {.Value = "0000"}},
                        New FontName() With {.Val = "Calibri"}))

        'Fills
        Dim FillIDs As New List(Of Integer)
        Dim myFill As Fills = mySSh.GetFirstChild(Of Fills)()
        If myFill Is Nothing Then
            mySSh.InsertAfter(New Fills(), myF)
            myFill = mySSh.GetFirstChild(Of Fills)()
        End If
        FillIDs.Add(myFill.ChildElements.Count)
        myFill.Append(New Fill(New PatternFill() With {.PatternType = PatternValues.None}))
        FillIDs.Add(myFill.ChildElements.Count)
        myFill.Append(New Fill(New PatternFill(New ForegroundColor With {.Theme = 3UI, .Tint = DoubleValue.FromDouble(0.59999389629810485D)},
                                                New BackgroundColor With {.Indexed = 64UI}) With {.PatternType = PatternValues.Solid}))
        FillIDs.Add(myFill.ChildElements.Count)
        myFill.Append(New Fill(New PatternFill(New ForegroundColor With {.Theme = 3UI, .Tint = DoubleValue.FromDouble(0.59999389629810485D)},
                                                New BackgroundColor With {.Indexed = 64UI}) With {.PatternType = PatternValues.Solid}))
        FillIDs.Add(myFill.ChildElements.Count)
        myFill.Append(New Fill(New PatternFill(New ForegroundColor With {.Theme = 8UI, .Tint = DoubleValue.FromDouble(0.39997558519241921D)},
                                                New BackgroundColor With {.Indexed = 64UI}) With {.PatternType = PatternValues.Solid}))
        FillIDs.Add(myFill.ChildElements.Count)
        myFill.Append(New Fill(New PatternFill(New ForegroundColor With {.Theme = 8UI, .Tint = DoubleValue.FromDouble(0.59999389629810485D)},
                                                New BackgroundColor With {.Indexed = 64UI}) With {.PatternType = PatternValues.Solid}))

        'Border
        Dim BorderID As Integer = 0
        Dim myBorders As Borders = mySSh.GetFirstChild(Of Borders)()
        If myBorders Is Nothing Then
            mySSh.InsertAfter(New Borders(), myFill)
            myBorders = mySSh.GetFirstChild(Of Borders)()
        End If
        BorderID = myBorders.ChildElements.Count
        myBorders.Append(New Border)

        'Styles
        Dim myReturn As New ExcelStyleDefs
        Dim myCellFormats As CellFormats = mySSh.GetFirstChild(Of CellFormats)()
        If myCellFormats Is Nothing Then
            mySSh.Append(New CellFormats())
            myCellFormats = mySSh.GetFirstChild(Of CellFormats)()
        End If
        myReturn.Add(CellFormatting.Null, myCellFormats.ChildElements.Count)
        myCellFormats.Append(New CellFormat() With {.FontId = FontIDs(0), .FillId = FillIDs(0), .BorderId = BorderID})
        myReturn.Add(CellFormatting.ColHeader, myCellFormats.ChildElements.Count)
        myCellFormats.Append(New CellFormat() With {.FontId = FontIDs(2), .FillId = FillIDs(2), .BorderId = BorderID, .ApplyFont = True, .ApplyFill = True})
        myReturn.Add(CellFormatting.RowHeader1, myCellFormats.ChildElements.Count)
        myCellFormats.Append(New CellFormat() With {.FontId = FontIDs(2), .FillId = FillIDs(3), .BorderId = BorderID, .ApplyFont = True, .ApplyFill = True})
        myReturn.Add(CellFormatting.RowHeader2, myCellFormats.ChildElements.Count)
        myCellFormats.Append(New CellFormat() With {.FontId = FontIDs(2), .FillId = FillIDs(4), .BorderId = BorderID, .ApplyFont = True, .ApplyFill = True})
        myReturn.Add(CellFormatting.NumInt, myCellFormats.ChildElements.Count)
        myCellFormats.Append(New CellFormat() With {.FontId = FontIDs(0), .FillId = FillIDs(0), .BorderId = BorderID, .NumberFormatId = 3UI, .ApplyNumberFormat = True})
        myReturn.Add(CellFormatting.Num2Dec, myCellFormats.ChildElements.Count)
        myCellFormats.Append(New CellFormat() With {.FontId = FontIDs(0), .FillId = FillIDs(0), .BorderId = BorderID, .NumberFormatId = 4UI, .ApplyNumberFormat = True})
        myReturn.Add(CellFormatting.Hyperlink, myCellFormats.ChildElements.Count)
        myCellFormats.Append(New CellFormat() With {.FontId = FontIDs(1), .FillId = FillIDs(0), .BorderId = BorderID, .ApplyFont = True})
        myReturn.Add(CellFormatting.MultiLineText, myCellFormats.ChildElements.Count)
        myCellFormats.Append(New CellFormat() With {.FontId = FontIDs(0), .FillId = FillIDs(0), .BorderId = BorderID, .ApplyAlignment = True, .Alignment = New Alignment With {.WrapText = True}})


        mySSh.Save()

        Return myReturn
    End Function

    Public Shared Sub SetCellFormula(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal WSP As WorksheetPart, ByVal FormulaExpression As String)
        Dim myCell As Cell = InsertCellInWorksheet(columnName, rowIndex, WSP, 0UI)
        If myCell IsNot Nothing Then
            myCell.CellFormula = New CellFormula() With {.FormulaType = CellFormulaValues.Normal, .Text = FormulaExpression}
            If myCell.CellValue IsNot Nothing Then myCell.CellValue.Remove()
        End If
    End Sub

    'Public Shared Sub FreezePanes(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal wsPart As WorksheetPart)
    '    Dim mySheetViews As SheetViews = wsPart.Worksheet.GetFirstChild(Of SheetViews)()
    '    If mySheetViews Is Nothing Then
    '        wsPart.Worksheet.Append(New SheetViews(
    '                                    New SheetView(
    '                                        New Pane With {.VerticalSplit = 1D, .TopLeftCell = columnName.ToUpper + rowIndex.ToString.Trim,
    '                                                        .ActivePane = PaneValues.BottomLeft, .State = PaneStateValues.Frozen}
    '                          )))
    '    Else
    '        Try
    '            Dim sv As SheetView = mySheetViews.GetFirstChild(Of SheetView)()
    '            Dim Sel As Selection = sv.GetFirstChild(Of Selection)()
    '            Dim P As New Pane() With {.VerticalSplit = 1D, .TopLeftCell = columnName.ToUpper + rowIndex.ToString.Trim,
    '                                      .ActivePane = PaneValues.BottomLeft, .State = PaneStateValues.Frozen}
    '            sv.InsertBefore(P, Sel)
    '            Sel.Pane = PaneValues.BottomLeft
    '        Catch ex As Exception
    '            'hm
    '        End Try
    '    End If
    'End Sub

    Public Shared Function GetWorksheetNameFromRefStr(RefStr As String) As String
        Dim myreturn As String = ""
        If Not String.IsNullOrEmpty(RefStr) Then
            Dim pos As Integer = RefStr.IndexOf("!$")
            If pos > 0 Then
                myreturn = RefStr.Substring(0, pos)
                If myreturn.Substring(0, 1) = "'" Then myreturn = myreturn.Substring(1, myreturn.Length - 2)
            End If
        End If

        Return myreturn
    End Function

    Public Shared Function GetColumnFromRefStr(RefStr As String) As String
        Dim myreturn As String = ""
        If Not String.IsNullOrEmpty(RefStr) Then
            Dim pos As Integer = RefStr.IndexOf("!$")
            If pos > 0 Then
                myreturn = RefStr.Substring(pos + 2) '$-Zeichen
                myreturn = myreturn.Substring(0, myreturn.IndexOf("$"))
            End If
        End If

        Return myreturn
    End Function

    Public Shared Function GetRowFromRefStr(RefStr As String) As Integer
        Dim myreturn As Integer = 0
        If Not String.IsNullOrEmpty(RefStr) Then
            Dim mc As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(RefStr, "!\$[A-Z]+\$[0-9]+")
            If mc.Count > 0 Then
                Dim tmpstr As String = mc.Item(0).Value
                tmpstr = tmpstr.Substring(tmpstr.LastIndexOf("$") + 1)
                Try
                    myreturn = Integer.Parse(tmpstr)
                Catch ex As Exception
                    myreturn = 0
                End Try
            End If
        End If

        Return myreturn
    End Function

    Public Shared Function GetCellValueFromRefStr(ByRef document As SpreadsheetDocument, RefStr As String) As String
        Dim myreturn As String
        Try
            myreturn = GetCellValue(document, GetWorksheetNameFromRefStr(RefStr), GetColumnFromRefStr(RefStr) + GetRowFromRefStr(RefStr).ToString)
        Catch ex As Exception
            myreturn = Nothing
        End Try
        Return myreturn
    End Function

    Public Shared Function GetDefinedNameValue(ByRef document As SpreadsheetDocument, NameName As String) As String
        Dim myreturn As String = ""
        Dim WB As Workbook = document.WorkbookPart.Workbook
        If WB.DefinedNames IsNot Nothing Then
            For Each dn As DefinedName In WB.DefinedNames
                If dn.Name.Value = NameName Then
                    myreturn = dn.Text
                    Exit For
                End If
            Next
        End If
        Return myreturn
    End Function

    Public Shared Sub SetCellValueHyperlink(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal WSP As WorksheetPart, CellContent As String,
                                            LinkAddress As String, Optional StyleDef As ExcelStyleDefs = Nothing)
        'If System.Uri.IsWellFormedUriString(LinkAddress, UriKind.Relative) Then
        Dim HLinks As Hyperlinks
        If WSP.Worksheet.Descendants(Of Hyperlinks)().Count = 0 Then
            HLinks = New Hyperlinks
            If WSP.Worksheet.Descendants(Of PageMargins)().Count = 0 Then
                WSP.Worksheet.InsertAfter(Of Hyperlinks)(HLinks, WSP.Worksheet.Descendants(Of SheetData)().First())
            Else
                WSP.Worksheet.InsertBefore(Of Hyperlinks)(HLinks, WSP.Worksheet.Descendants(Of PageMargins)().First())
            End If
        Else
            HLinks = WSP.Worksheet.Descendants(Of Hyperlinks)().First
        End If

        Dim HLinkId As String = "ryId" + (HLinks.ChildElements.Count + 1).ToString
        HLinks.Append(New Hyperlink With {.Reference = columnName + rowIndex.ToString().Trim, .Id = HLinkId})

        WSP.AddHyperlinkRelationship(New System.Uri(LinkAddress, UriKind.Relative), True, HLinkId)
        WSP.Worksheet.Save()

        SetCellValueString(columnName, rowIndex, WSP, CellContent, CellFormatting.Hyperlink, StyleDef)
        'Else
        '    SetCellValueString(columnName, rowIndex, WSP, CellContent, CellFormatting.Null, StyleDef)
        'End If
    End Sub

    '#########################################################################
    ''' <summary>
    ''' Öffnet xlsx in RO-Modus und liest komplett in Speicher; kein Schließen des Dokumentes nötig;
    ''' wirft alle möglichen Exceptions
    ''' </summary>
    Public Shared Function OpenSpreadSheetReadOnly(filename As String) As SpreadsheetDocument
        Return SpreadsheetDocument.Open(New IO.MemoryStream(IO.File.ReadAllBytes(filename)), False)
    End Function


    '#########################################################################
    Public Shared Function GetCellValueDate(ByRef document As SpreadsheetDocument, ByVal sheetName As String, ByVal addressName As String) As DateTime
        Dim value As String = GetCellValue(document, sheetName, addressName)
        Dim myreturn As DateTime = DateTime.MinValue

        If value IsNot Nothing Then
            Try
                myreturn = DateTime.FromOADate(Convert.ToDouble(value))
            Catch ex As Exception
                myreturn = DateTime.MinValue
            End Try
        End If

        Return myreturn
    End Function

    '#########################################################################
    Public Shared Function GetWorksheetPart(ByRef document As SpreadsheetDocument, ByVal sheetName As String) As WorksheetPart
        Dim myreturn As WorksheetPart = Nothing
        ' Retrieve a reference to the workbook part.
        Dim wbPart As WorkbookPart = document.WorkbookPart

        ' Find the sheet with the supplied name, and then use that Sheet object
        ' to retrieve a reference to the appropriate worksheet.
        Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
            Where(Function(s) s.Name = sheetName).FirstOrDefault()

        If theSheet IsNot Nothing Then myreturn = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)

        Return myreturn
    End Function

    '#########################################################################
    Public Shared Function GetCellValue(ByRef document As SpreadsheetDocument, ByVal sheetName As String, ByVal addressName As String) As String
        Dim value As String = Nothing

        ' Retrieve a reference to the workbook part.
        Dim wbPart As WorkbookPart = document.WorkbookPart

        ' Find the sheet with the supplied name, and then use that Sheet object
        ' to retrieve a reference to the appropriate worksheet.
        Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
            Where(Function(s) s.Name = sheetName).FirstOrDefault()

        If theSheet IsNot Nothing Then
            Dim wsPart As WorksheetPart = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)

            ' Use its Worksheet property to get a reference to the cell 
            ' whose address matches the address you supplied.
            Dim theCell As Cell = wsPart.Worksheet.Descendants(Of Cell).
                Where(Function(c) c.CellReference = addressName).FirstOrDefault

            ' If the cell does not exist, return an empty string.
            If theCell IsNot Nothing Then
                If theCell.CellValue Is Nothing Then
                    value = theCell.InnerText
                Else
                    value = theCell.CellValue.InnerText
                End If


                ' If the cell represents an numeric value, you are done. 
                ' For dates, this code returns the serialized value that 
                ' represents the date. The code handles strings and 
                ' Booleans individually. For shared strings, the code 
                ' looks up the corresponding value in the shared string 
                ' table. For Booleans, the code converts the value into 
                ' the words TRUE or FALSE.
                If theCell.DataType IsNot Nothing Then
                    Select Case theCell.DataType.Value
                        Case CellValues.SharedString

                            ' For shared strings, look up the value in the 
                            ' shared strings table.
                            Dim stringTable = wbPart.
                                GetPartsOfType(Of SharedStringTablePart).FirstOrDefault()

                            ' If the shared string table is missing, something
                            ' is wrong. Return the index that is in 
                            ' the cell. Otherwise, look up the correct text in 
                            ' the table.
                            If stringTable IsNot Nothing Then
                                value = stringTable.SharedStringTable.
                                ElementAt(Integer.Parse(value)).InnerText
                            End If

                        Case CellValues.Boolean
                            Select Case value
                                Case "0"
                                    value = "FALSE"
                                Case Else
                                    value = "TRUE"
                            End Select
                    End Select
                End If
            End If
        End If

        Return value
    End Function

    '#########################################################################
    Public Shared Function HasCellFormula(ByRef document As SpreadsheetDocument, ByVal sheetName As String, ByVal addressName As String) As Boolean
        ' Retrieve a reference to the workbook part.
        Dim wbPart As WorkbookPart = document.WorkbookPart

        ' Find the sheet with the supplied name, and then use that Sheet object
        ' to retrieve a reference to the appropriate worksheet.
        Dim theSheet As Sheet = wbPart.Workbook.Descendants(Of Sheet)().
            Where(Function(s) s.Name = sheetName).FirstOrDefault()

        If theSheet IsNot Nothing Then
            Dim wsPart As WorksheetPart = CType(wbPart.GetPartById(theSheet.Id), WorksheetPart)

            ' Use its Worksheet property to get a reference to the cell 
            ' whose address matches the address you supplied.
            Dim theCell As Cell = wsPart.Worksheet.Descendants(Of Cell).
                Where(Function(c) c.CellReference = addressName).FirstOrDefault

            ' If the cell does not exist, return an empty string.
            If theCell IsNot Nothing Then
                Return theCell.CellFormula IsNot Nothing
            End If
        End If
        Return False
    End Function

    '#########################################################################
    ''' <summary>
    ''' Riskant: keine Prüfungen, sondern gleich Anhängen; ACHTUNG: kein worksheetPart.Worksheet.Save() !!
    ''' </summary>
    Public Shared Sub AppendRow(ByVal rowIndex As UInteger, ByRef RowDataList As List(Of RowData), ByVal worksheetPart As WorksheetPart, Optional StyleDef As ExcelStyleDefs = Nothing)
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim row As New Row With {.RowIndex = rowIndex}

        If StyleDef Is Nothing Then StyleDef = xlsxFactory.StandardExcelStyleDefs

        For Each c As RowData In (From rd As RowData In RowDataList Where Not String.IsNullOrEmpty(rd.Column.Trim) Let sortstr = rd.Column.PadLeft(6) Order By sortstr Select rd).ToList

            If c.CellType = CellTypes.int Then
                row.InsertBefore(New Cell With {.CellReference = c.Column + rowIndex.ToString().Trim,
                                                .StyleIndex = StyleDef(CellFormatting.NumInt),
                                                .CellValue = New CellValue(c.Value),
                                                .DataType = New EnumValue(Of CellValues)(CellValues.Number)}, Nothing)
            ElseIf c.CellType = CellTypes.dec Then
                Dim myCellValue As Double = 0
                If Double.TryParse(c.Value, myCellValue) Then
                    c.Value = myCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture)
                Else
                    c.Value = "0"
                End If
                row.InsertBefore(New Cell With {.CellReference = c.Column + rowIndex.ToString().Trim,
                                                .StyleIndex = StyleDef(CellFormatting.Num2Dec),
                                                .CellValue = New CellValue(c.Value),
                                                .DataType = New EnumValue(Of CellValues)(CellValues.Number)}, Nothing)
            ElseIf c.CellType = CellTypes.hyperlink AndAlso System.Uri.IsWellFormedUriString(c.Link, UriKind.Relative) Then
                '#######
                Dim HLinks As Hyperlinks
                If worksheetPart.Worksheet.Descendants(Of Hyperlinks)().Count = 0 Then
                    HLinks = New Hyperlinks
                    If worksheetPart.Worksheet.Descendants(Of PageMargins)().Count = 0 Then
                        worksheetPart.Worksheet.InsertAfter(Of Hyperlinks)(HLinks, worksheetPart.Worksheet.Descendants(Of SheetData)().First())
                    Else
                        worksheetPart.Worksheet.InsertBefore(Of Hyperlinks)(HLinks, worksheetPart.Worksheet.Descendants(Of PageMargins)().First())
                    End If
                Else
                    HLinks = worksheetPart.Worksheet.Descendants(Of Hyperlinks)().First
                End If

                Dim HLinkId As String = "rId" + (HLinks.ChildElements.Count + 1).ToString
                HLinks.Append(New Hyperlink With {.Reference = c.Column + rowIndex.ToString().Trim, .Id = HLinkId})

                worksheetPart.Worksheet.Save()
                worksheetPart.AddHyperlinkRelationship(New System.Uri(c.Link, UriKind.Relative), True, HLinkId)

                row.InsertBefore(New Cell With {.CellReference = c.Column + rowIndex.ToString().Trim,
                                                .StyleIndex = StyleDef(CellFormatting.Hyperlink),
                                                .CellValue = New CellValue(c.Value),
                                                .DataType = New EnumValue(Of CellValues)(CellValues.String)}, Nothing)
                '#######
            ElseIf c.CellType = CellTypes.datetime Then
                row.InsertBefore(New Cell With {.CellReference = c.Column + rowIndex.ToString().Trim,
                                                .StyleIndex = StyleDef(CellFormatting.Null),
                                                .CellValue = New CellValue(c.Value),
                                                .DataType = New EnumValue(Of CellValues)(CellValues.String)}, Nothing)
            ElseIf c.CellType = CellTypes.text Then
                row.InsertBefore(New Cell With {.CellReference = c.Column + rowIndex.ToString().Trim,
                                                .StyleIndex = StyleDef(CellFormatting.MultiLineText),
                                                .CellValue = New CellValue(c.Value),
                                                .DataType = New EnumValue(Of CellValues)(CellValues.String)}, Nothing)
            ElseIf c.CellType = CellTypes.formula Then
                row.InsertBefore(New Cell With {.CellReference = c.Column + rowIndex.ToString().Trim,
                                                .StyleIndex = StyleDef(CellFormatting.NumInt),
                                                .CellValue = Nothing,
                                                .CellFormula = New CellFormula() With {.FormulaType = CellFormulaValues.Normal, .Text = c.Value},
                                                .DataType = New EnumValue(Of CellValues)(CellValues.Number)}, Nothing)

            Else 'CellTypes.str
                row.InsertBefore(New Cell With {.CellReference = c.Column + rowIndex.ToString().Trim,
                                                .StyleIndex = StyleDef(CellFormatting.Null),
                                                .CellValue = New CellValue(c.Value),
                                                .DataType = New EnumValue(Of CellValues)(CellValues.String)}, Nothing)
            End If
        Next

        sheetData.Append(row)
    End Sub


    Public Shared Sub DefineName(ByRef document As SpreadsheetDocument, NewName As String, NewValue As String)
        Dim WB As Workbook = document.WorkbookPart.Workbook
        If WB.DefinedNames Is Nothing Then WB.DefinedNames = New DefinedNames
        For Each dn As DefinedName In WB.DefinedNames
            If dn.Name.Value = NewName Then
                dn.Text = NewValue
                WB.Save()
                Return
            End If
        Next
        WB.DefinedNames.Append(New DefinedName With {.Name = NewName, .Text = NewValue})
        WB.Save()
    End Sub

    Public Shared Sub SetColumnWidth(ByVal columnName As String, ByVal WSP As WorksheetPart, ColumnWidth As Double)
        Dim SpreadSheetColumns As Columns
        If WSP.Worksheet.Descendants(Of Columns)().Count = 0 Then
            SpreadSheetColumns = New Columns
            If WSP.Worksheet.Descendants(Of SheetData)().Count = 0 Then
                'Fehler: SheetData muss es geben!
                Return
            Else
                WSP.Worksheet.InsertBefore(Of Columns)(SpreadSheetColumns, WSP.Worksheet.Descendants(Of SheetData)().First())
            End If
        Else
            SpreadSheetColumns = WSP.Worksheet.Descendants(Of Columns)().First
        End If
        Dim colIndex As Integer = GetColumnIndexFromString(columnName)
        SpreadSheetColumns.Append(New Column With {.CustomWidth = True, .Min = colIndex, .Max = colIndex, .Width = ColumnWidth})
    End Sub

    Private Shared Function GetColumnIndexFromString(ColStr As String) As Integer
        ColStr = ColStr.ToUpper
        Dim myreturn As Integer = Asc(ColStr.Substring(ColStr.Length - 1, 1)) - Asc("A") + 1
        If ColStr.Length > 1 Then
            Dim multiplier As Integer = Asc("Z") - Asc("A") + 1
            myreturn = myreturn + (Asc(ColStr.Substring(ColStr.Length - 2, 1)) - Asc("A") + 1) * multiplier
            If ColStr.Length > 2 Then
                myreturn = myreturn + (Asc(ColStr.Substring(ColStr.Length - 3, 1)) - Asc("A") + 1) * multiplier * multiplier
            End If
        End If

        Return myreturn
    End Function

    Private Shared Function InsertSharedStringItem(sharedString As String, sharedStringTable As SharedStringTable, sharedStrings As Dictionary(Of String, Integer)) As Integer
        Dim sharedStringIndex As Integer

        If Not sharedStrings.TryGetValue(sharedString, sharedStringIndex) Then
            'The text does not exist in the part. Create the SharedStringItem now.
            sharedStringTable.AppendChild(New SharedStringItem(New Text(sharedString)))
            sharedStringIndex = sharedStrings.Count
            sharedStrings.Add(sharedString, sharedStringIndex)
        End If

        Return sharedStringIndex
    End Function

End Class

Public Class RowData
    Public Column As String
    Public Value As String
    Public CellType As CellTypes
    Public Link As String
End Class
