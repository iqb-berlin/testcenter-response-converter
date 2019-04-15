Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging

Public Class ConvertResponses

    Public Shared Sub Run(myworker As ComponentModel.BackgroundWorker, sourceFolder As String)
        Dim targetXlsxFilename As String = sourceFolder + IO.Path.DirectorySeparatorChar + "logs " + IO.Path.GetFileName(sourceFolder) + ".xlsx"
        Dim myTemplate As Byte() = Nothing
        Try
            Dim TmpZielXLS As SpreadsheetDocument = SpreadsheetDocument.Create(targetXlsxFilename, SpreadsheetDocumentType.Workbook)
            Dim myWorkbookPart As WorkbookPart = TmpZielXLS.AddWorkbookPart()
            myWorkbookPart.Workbook = New Workbook()
            myWorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())
            TmpZielXLS.Close()

            myTemplate = IO.File.ReadAllBytes(targetXlsxFilename)
        Catch ex As Exception
            myworker.ReportProgress(20.0#, "e: Konnte Datei '" + targetXlsxFilename + "' nicht schreiben (noch geöffnet?)" + vbNewLine + ex.Message)
        End Try

        If myTemplate IsNot Nothing Then
            Const LogFileFirstLine = "groupname;loginname;code;bookletname;unitname;timestamp;logentry"

            Dim LogStarts As New Dictionary(Of String, Long) 'id -> time
            Dim FirstBookletLoadStarts As New Dictionary(Of String, Long) 'id -> time
            Dim FirstBookletLoadCompletes As New Dictionary(Of String, Long) 'id -> time
            Dim FirstUnitEnter As New Dictionary(Of String, Long) 'id -> time
            Dim Events As New List(Of String)
            'Dim LogData As New Dictionary(Of String, Dictionary(Of String, Long))
            Dim SearchDir As New IO.DirectoryInfo(sourceFolder)
            For Each fi As IO.FileInfo In SearchDir.GetFiles("*.csv", IO.SearchOption.AllDirectories)
                If myworker.CancellationPending Then Exit For

                Dim line As String
                Dim readFile As System.IO.TextReader = New IO.StreamReader(fi.FullName)
                Dim isLogFile As Boolean
                Try
                    line = readFile.ReadLine()
                    isLogFile = line = LogFileFirstLine
                Catch ex As Exception
                    line = ""
                    isLogFile = False
                End Try
                If isLogFile Then
                    myworker.ReportProgress(20.0#, "Lese " + fi.Name)
                    Do While line IsNot Nothing
                        line = readFile.ReadLine()
                        If line IsNot Nothing Then
                            '######
                            Dim lineSplits As String() = line.Split({""";"}, StringSplitOptions.RemoveEmptyEntries)
                            If lineSplits.Count = 7 Then
                                Dim group As String = lineSplits(0).Substring(1)
                                Dim login As String = lineSplits(1).Substring(1)
                                Dim code As String = lineSplits(2).Substring(1)
                                Dim booklet As String = lineSplits(3).Substring(1)
                                Dim unit As String = lineSplits(4)
                                If unit.Length < 2 Then
                                    unit = ""
                                Else
                                    unit = unit.Substring(1)
                                End If
                                Dim timestampStr As String = lineSplits(5).Substring(1)
                                Dim timestampInt As Long = Long.Parse(timestampStr)
                                Dim entry As String = lineSplits(6)
                                Dim key As String = entry
                                If key.IndexOf(": ") > 0 Then key = key.Substring(0, key.IndexOf(": "))

                                Dim PersonKey As String = group + login + code
                                If Not LogStarts.ContainsKey(PersonKey) Then
                                    LogStarts.Add(PersonKey, timestampInt)
                                Else
                                    If LogStarts.Item(PersonKey) > timestampInt Then LogStarts.Item(PersonKey) = timestampInt
                                End If

                                Select Case key
                                    Case "UNITENTER"
                                        If Not FirstUnitEnter.ContainsKey(PersonKey) Then
                                            FirstUnitEnter.Add(PersonKey, timestampInt)
                                        Else
                                            If FirstUnitEnter.Item(PersonKey) > timestampInt Then FirstUnitEnter.Item(PersonKey) = timestampInt
                                        End If
                                    Case "BOOKLETLOADSTART"
                                        If Not FirstBookletLoadStarts.ContainsKey(PersonKey) Then
                                            FirstBookletLoadStarts.Add(PersonKey, timestampInt)
                                        Else
                                            If FirstBookletLoadStarts.Item(PersonKey) > timestampInt Then FirstBookletLoadStarts.Item(PersonKey) = timestampInt
                                        End If

                                    Case "BOOKLETLOADCOMPLETE"
                                        If Not FirstBookletLoadCompletes.ContainsKey(PersonKey) Then
                                            FirstBookletLoadCompletes.Add(PersonKey, timestampInt)
                                        Else
                                            If FirstBookletLoadCompletes.Item(PersonKey) > timestampInt Then FirstBookletLoadCompletes.Item(PersonKey) = timestampInt
                                        End If

                                    Case Else

                                End Select
                            End If
                            '######
                        End If
                    Loop
                End If
                readFile.Close()
                readFile = Nothing
            Next

            If Not myworker.CancellationPending Then

                Using MemStream As New IO.MemoryStream()
                    MemStream.Write(myTemplate, 0, myTemplate.Length)
                    Using ZielXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, True)
                        Dim myStyles As ExcelStyleDefs = xlsxFactory.AddIQBStandardStyles(ZielXLS.WorkbookPart)
                        Dim Tabelle_Classes As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Logs")

                        '########################################################
                        Dim currentUser As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent
                        Dim currentUserName As String = currentUser.Name.Substring(currentUser.Name.IndexOf("\") + 1)

                        xlsxFactory.SetCellValueString("A", 1, Tabelle_Classes, "Zeitpunkt-Variablen Testcenter", xlsxFactory.CellFormatting.Null, myStyles)
                        xlsxFactory.SetCellValueString("A", 2, Tabelle_Classes, "generiert mit " + My.Application.Info.AssemblyName + " V" +
                                                       My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString + "." +
                                                       My.Application.Info.Version.Build.ToString + " am " + DateTime.Now.ToShortDateString + " " + DateTime.Now.ToShortTimeString +
                                                       " (" + currentUserName + ")", xlsxFactory.CellFormatting.Null, myStyles)

                        Dim myRow As Integer = 4

                        xlsxFactory.SetCellValueString("A", myRow, Tabelle_Classes, "ID", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", Tabelle_Classes, 20)
                        xlsxFactory.SetCellValueString("B", myRow, Tabelle_Classes, "Start", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", Tabelle_Classes, 20)
                        xlsxFactory.SetCellValueString("C", myRow, Tabelle_Classes, "BookletLoadingStart", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("C", Tabelle_Classes, 20)
                        xlsxFactory.SetCellValueString("D", myRow, Tabelle_Classes, "BookletLoadingComplete", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("D", Tabelle_Classes, 20)
                        xlsxFactory.SetCellValueString("E", myRow, Tabelle_Classes, "FirstUnitEnter", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("E", Tabelle_Classes, 20)


                        For Each tc As KeyValuePair(Of String, Long) In LogStarts
                            If myworker.CancellationPending Then Exit For
                            myRow += 1
                            Dim myRowData As New List(Of RowData)
                            myRowData.Add(New RowData With {.Column = "A", .Value = tc.Key, .CellType = CellTypes.str})
                            myRowData.Add(New RowData With {.Column = "B", .Value = tc.Value, .CellType = CellTypes.int})
                            If FirstBookletLoadStarts.ContainsKey(tc.Key) Then
                                myRowData.Add(New RowData With {.Column = "C", .Value = FirstBookletLoadStarts.Item(tc.Key) - tc.Value, .CellType = CellTypes.int})
                            End If
                            If FirstBookletLoadCompletes.ContainsKey(tc.Key) Then
                                myRowData.Add(New RowData With {.Column = "D", .Value = FirstBookletLoadCompletes.Item(tc.Key) - tc.Value, .CellType = CellTypes.int})
                            End If
                            If FirstUnitEnter.ContainsKey(tc.Key) Then
                                myRowData.Add(New RowData With {.Column = "E", .Value = FirstUnitEnter.Item(tc.Key) - tc.Value, .CellType = CellTypes.int})
                            End If
                            xlsxFactory.AppendRow(myRow, myRowData, Tabelle_Classes)
                        Next


                    End Using
                    Try
                        Using fs As New IO.FileStream(targetXlsxFilename, IO.FileMode.Create)
                            MemStream.WriteTo(fs)
                        End Using
                        myworker.ReportProgress(20.0#, "Log-Einträge gespeichert")
                    Catch ex As Exception
                        myworker.ReportProgress(20.0#, "e: Konnte Datei nicht schreiben: " + ex.Message)
                    End Try
                End Using
            End If
        End If
    End Sub

End Class

