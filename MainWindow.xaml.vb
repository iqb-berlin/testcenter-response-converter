Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports Newtonsoft.Json

Class MainWindow
    Private bookletSize As New Dictionary(Of String, Long)
    Const LogFileFirstLine = "groupname;loginname;code;bookletname;unitname;timestamp;logentry"
    Const ResponsesFileFirstLine = "groupname;loginname;code;bookletname;unitname;responses;restorePoint;responseType;response-ts;restorePoint-ts;laststate"

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        Me.Title = My.Application.Info.AssemblyName

    End Sub

    Private Sub BtnChangeCsvSource_Click(sender As Object, e As RoutedEventArgs)
        Dim folderpicker As New System.Windows.Forms.FolderBrowserDialog With {.Description = "Wählen des Quellverzeichnisses für die Csv-Dateien",
                                                        .ShowNewFolderButton = False, .SelectedPath = My.Settings.lastfolder_csvSources}
        If folderpicker.ShowDialog() Then
            Me.TBCsvSource.Text = folderpicker.SelectedPath
            My.Settings.lastfolder_csvSources = folderpicker.SelectedPath
            My.Settings.Save()

            Dim SearchDir As New IO.DirectoryInfo(My.Settings.lastfolder_csvSources)
            Dim LogFileCount As Integer = 0
            Dim ResponsesFileCount As Integer = 0
            For Each fi As IO.FileInfo In SearchDir.GetFiles("*.csv", IO.SearchOption.AllDirectories)
                Try
                    Dim readFile As System.IO.TextReader = New IO.StreamReader(fi.FullName)
                    Dim line As String = readFile.ReadLine()
                    If line = LogFileFirstLine Then
                        LogFileCount += 1
                    ElseIf line = ResponsesFileFirstLine Then
                        ResponsesFileCount += 1
                    Else
                        Me.MBUC.AddMessage("w: Datei nicht erkannt: " + fi.Name)
                    End If
                Catch ex As Exception
                    Me.MBUC.AddMessage("e: Fehler beim Lesen der Datei " + fi.Name + "; noch geöffnet?")
                End Try
            Next
            Me.MBUC.AddMessage(LogFileCount.ToString + " Log-Dateien und " + ResponsesFileCount.ToString + " Antwortdateien erkannt.")
        End If
    End Sub

    Private Sub BtnChangeBookletTxtSource_Click(sender As Object, e As RoutedEventArgs)
        Dim filepicker As New Microsoft.Win32.OpenFileDialog With {.FileName = My.Settings.lastfile_BookletTxt, .Filter = "Txt-Dateien|*.txt",
                                                                           .DefaultExt = "Xlsx", .Title = "BookletTxt - Wähle Datei"}
        If filepicker.ShowDialog Then
            Me.TBBookletTxtSource.Text = filepicker.FileName
            My.Settings.lastfile_BookletTxt = filepicker.FileName
            My.Settings.Save()

            Me.MBUC.AddMessage("Lese Bookletdatei")
            Dim bookletline As String
            Dim readFile As System.IO.TextReader = New IO.StreamReader(My.Settings.lastfile_BookletTxt)
            Try
                bookletline = readFile.ReadLine()
            Catch ex As Exception
                bookletline = ""
                Me.MBUC.AddMessage("e:Fehler beim Lesen der Bookletdatei: " + ex.Message)
            End Try
            If Not String.IsNullOrEmpty(bookletline) Then
                bookletSize.Clear()
                Do While bookletline IsNot Nothing
                    Dim lineSplits As String() = bookletline.Split({" "}, StringSplitOptions.RemoveEmptyEntries)
                    If lineSplits.Count = 2 Then
                        Dim tryInt As Integer = 0
                        If Long.TryParse(lineSplits(1), tryInt) AndAlso Not bookletSize.ContainsKey(lineSplits(0).ToUpper()) Then
                            bookletSize.Add(lineSplits(0).ToUpper(), tryInt)
                        End If
                    End If
                    bookletline = readFile.ReadLine()
                Loop
                Me.MBUC.AddMessage(bookletSize.Count.ToString + " Einträge für Booklet-Größe gelesen")
            Else
                Me.MBUC.AddMessage("e:Bookletdatei ist leer")
            End If
        End If
    End Sub
    Private Sub BtnHelp_Click(sender As Object, e As RoutedEventArgs)
        MsgBox("Diese Funktion ist noch nicht verfügbar.", vbOK, "Warte warte, bald")
    End Sub

    Private Sub BtnEditor_Click() Handles BtnEditor.Click
        MBUC.StartEditorWithText()
    End Sub


    '##################################################################################
#Region "BackgroundWorker"
    '##################################################################################
    Private WithEvents Main_bw As ComponentModel.BackgroundWorker = Nothing
    Private Sub BtnCancel_Click() Handles BtnCancel.Click
        If Main_bw IsNot Nothing AndAlso Main_bw.IsBusy Then
            Main_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        End If
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Main_bw.ProgressChanged
        If e.ProgressPercentage > 0 AndAlso e.ProgressPercentage <= 100 Then
            Me.PB.Value = e.ProgressPercentage
            Me.PB.IsIndeterminate = False
        Else
            Me.PB.IsIndeterminate = True
        End If
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Main_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Main_bw.RunWorkerCompleted
        Me.PB.Value = 0.0#
        Me.PB.IsIndeterminate = False
        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")
        MBUC.AddMessage("beendet")
        BtnCancel.IsEnabled = False
    End Sub

    Private Sub BtnStart_Click(sender As Object, e As RoutedEventArgs)
        BtnCancel.IsEnabled = True
        Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_targetXlsx, .Filter = "Excel-Dateien|*.xlsx",
                                                           .DefaultExt = "xlsx", .Title = "Antworten Zieldatei wählen"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_targetXlsx = filepicker.FileName
            My.Settings.Save()

            Main_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            Main_bw.RunWorkerAsync()
        End If
    End Sub

#End Region
    '##################################################################################
    Private Sub Main_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Main_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender
        Dim targetXlsxFilename As String = My.Settings.lastfile_targetXlsx
        Dim myTemplate As Byte() = Nothing
        Try
            Dim TmpZielXLS As SpreadsheetDocument = SpreadsheetDocument.Create(targetXlsxFilename, SpreadsheetDocumentType.Workbook)
            Dim myWorkbookPart As WorkbookPart = TmpZielXLS.AddWorkbookPart()
            myWorkbookPart.Workbook = New Workbook()
            myWorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())
            TmpZielXLS.Close()

            myTemplate = IO.File.ReadAllBytes(targetXlsxFilename)
        Catch ex As Exception
            myworker.ReportProgress(0.0#, "e: Konnte Datei '" + targetXlsxFilename + "' nicht schreiben (noch geöffnet?)" + vbNewLine + ex.Message)
        End Try

        If myTemplate IsNot Nothing Then
            Dim myTestPersonList As New TestPersonList
            Dim Events As New List(Of String)
            Dim AllData As New SortedDictionary(Of String, Dictionary(Of String, List(Of ResponseEntry))) 'id -> booklet -> entries
            Dim AllVariables As New List(Of String)
            Dim AllUnitsWithResponses As New List(Of String)
            Dim LogEntryCount As Long = 0

            'Dim LogData As New Dictionary(Of String, Dictionary(Of String, Long))
            Dim SearchDir As New IO.DirectoryInfo(My.Settings.lastfolder_csvSources)
            For Each fi As IO.FileInfo In SearchDir.GetFiles("*.csv", IO.SearchOption.AllDirectories)
                If myworker.CancellationPending Then Exit For

                Dim line As String
                Dim readFile As System.IO.TextReader = Nothing
                Try
                    readFile = New IO.StreamReader(fi.FullName)
                    line = readFile.ReadLine()
                Catch ex As Exception
                    line = ""
                    readFile = Nothing
                    myworker.ReportProgress(0.0#, "e:Fehler mein Lesen von " + fi.Name + "; noch geöffnet?")
                End Try
                If readFile IsNot Nothing Then
                    myworker.ReportProgress(0.0#, "Lese " + fi.Name)
                    If line = LogFileFirstLine Then
                        '#########################
                        Do While line IsNot Nothing
                            line = readFile.ReadLine()
                            If line IsNot Nothing Then
                                Dim lineSplits As String() = line.Split({""";"}, StringSplitOptions.RemoveEmptyEntries)
                                If lineSplits.Count = 7 Then
                                    LogEntryCount += 1
                                    Dim group As String = lineSplits(0).Substring(1)
                                    Dim login As String = lineSplits(1).Substring(1)
                                    Dim code As String = lineSplits(2).Substring(1)
                                    Dim booklet As String = lineSplits(3).Substring(1).ToUpper()
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
                                    Dim parameter As String = ""
                                    If key.IndexOf(": ") > 0 Then
                                        parameter = key.Substring(key.IndexOf(": ") + 2)
                                        key = key.Substring(0, key.IndexOf(": "))
                                    End If

                                    Select Case key
                                        Case "BOOKLETLOADSTART"
                                            myTestPersonList.SetFirstBookletLoadStart(group, login, code, booklet, timestampInt)

                                            Dim parameterClean As String = parameter.Replace("""""", """")
                                            parameterClean = parameterClean.Substring(1, parameterClean.Length - 2)
                                            Dim sysdata As Dictionary(Of String, String) = Nothing
                                            Try
                                                sysdata = JsonConvert.DeserializeObject(parameterClean, GetType(Dictionary(Of String, String)))
                                            Catch ex As Exception
                                                sysdata = Nothing
                                                Debug.Print("sysdata json convert failed: " + ex.Message)
                                            End Try
                                            myTestPersonList.SetSysdata(group, login, code, booklet, sysdata)
                                        Case "BOOKLETLOADCOMPLETE"
                                            myTestPersonList.SetFirstBookletLoadComplete(group, login, code, booklet, timestampInt)
                                        Case "RESPONSESCOMPLETE", "PRESENTATIONCOMPLETE"
                                            'ignore

                                        Case Else
                                            If Not String.IsNullOrEmpty(unit) Then
                                                myTestPersonList.AddUnitEvent(group, login, code, booklet, timestampInt, unit, key, parameter)
                                            End If
                                    End Select
                                End If
                            End If
                        Loop
                    ElseIf line = ResponsesFileFirstLine Then
                        '#########################
                        Dim lineCount As Integer = 1
                        Do While line IsNot Nothing
                            line = readFile.ReadLine()
                            lineCount += 1
                            If line IsNot Nothing Then
                                '######
                                Dim entry As New ResponseEntry(line, "file '" + fi.Name + "', line " + lineCount.ToString())
                                If entry.data.Count > 0 Then
                                    For Each d As KeyValuePair(Of String, String) In entry.data
                                        If Not AllUnitsWithResponses.Contains(entry.unit) Then AllUnitsWithResponses.Add(entry.unit)
                                        If Not AllVariables.Contains(entry.unit + "##" + d.Key) Then AllVariables.Add(entry.unit + "##" + d.Key)
                                    Next

                                    If Not AllData.ContainsKey(entry.Key) Then AllData.Add(entry.Key, New Dictionary(Of String, List(Of ResponseEntry)))
                                    Dim myPerson As Dictionary(Of String, List(Of ResponseEntry)) = AllData.Item(entry.Key)
                                    If Not myPerson.ContainsKey(entry.booklet) Then myPerson.Add(entry.booklet, New List(Of ResponseEntry))
                                    myPerson.Item(entry.booklet).Add(entry)
                                End If
                            End If
                            '######
                        Loop
                    End If
                    readFile.Close()
                End If
            Next

            myworker.ReportProgress(0.0#, "Daten für " + AllData.Count.ToString("#,##0") + " Testpersonen und " + AllVariables.Count.ToString("#,##0") + " Variablen gelesen.")
            myworker.ReportProgress(0.0#, LogEntryCount.ToString("#,##0") + " Log-Einträge gelesen.")

            If Not myworker.CancellationPending Then

                Using MemStream As New IO.MemoryStream()
                    MemStream.Write(myTemplate, 0, myTemplate.Length)
                    Using ZielXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, True)
                        Dim myStyles As ExcelStyleDefs = xlsxFactory.AddIQBStandardStyles(ZielXLS.WorkbookPart)
                        '########################################################
                        'Responses
                        '########################################################
                        Dim TableResponses As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Responses")

                        '########################################################
                        Dim currentUser As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent
                        Dim currentUserName As String = currentUser.Name.Substring(currentUser.Name.IndexOf("\") + 1)

                        xlsxFactory.SetCellValueString("A", 1, TableResponses, "Antwort-Variablen Testcenter", xlsxFactory.CellFormatting.Null, myStyles)
                        xlsxFactory.SetCellValueString("A", 2, TableResponses, "generiert mit " + My.Application.Info.AssemblyName + " V" +
                                                       My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString + "." +
                                                       My.Application.Info.Version.Build.ToString + " am " + DateTime.Now.ToShortDateString + " " + DateTime.Now.ToShortTimeString +
                                                       " (" + currentUserName + ")", xlsxFactory.CellFormatting.Null, myStyles)

                        Dim myRow As Integer = 4

                        xlsxFactory.SetCellValueString("A", myRow, TableResponses, "ID", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableResponses, 20)
                        xlsxFactory.SetCellValueString("B", myRow, TableResponses, "Booklet", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", TableResponses, 20)

                        Dim myColumn As String = "C"
                        Dim Columns As New Dictionary(Of String, String)

                        For Each s As String In From v As String In AllVariables Order By v Select v
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableResponses, s, xlsxFactory.CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableResponses, 10)
                            Columns.Add(s, myColumn)
                            myColumn = xlsxFactory.GetNextColumn(myColumn)
                        Next

                        Dim progressMax As Integer = AllData.Count
                        Dim progressCount As Integer = 1
                        For Each persondata As KeyValuePair(Of String, Dictionary(Of String, List(Of ResponseEntry))) In AllData
                            If myworker.CancellationPending Then Exit For
                            myworker.ReportProgress(progressCount * 100 / progressMax, "")
                            progressCount += 1
                            For Each bookletdata As KeyValuePair(Of String, List(Of ResponseEntry)) In persondata.Value
                                myRow += 1
                                Dim myRowData As New List(Of RowData)
                                If persondata.Value.Count > 1 Then
                                    myRowData.Add(New RowData With {.Column = "A", .Value = persondata.Key + bookletdata.Key, .CellType = CellTypes.str})
                                Else
                                    myRowData.Add(New RowData With {.Column = "A", .Value = persondata.Key, .CellType = CellTypes.str})
                                End If
                                myRowData.Add(New RowData With {.Column = "B", .Value = bookletdata.Key, .CellType = CellTypes.str})
                                For Each u As ResponseEntry In bookletdata.Value
                                    For Each d As KeyValuePair(Of String, String) In u.data
                                        myRowData.Add(New RowData With {.Column = Columns.Item(u.unit + "##" + d.Key), .Value = d.Value, .CellType = CellTypes.str})
                                    Next
                                Next
                                xlsxFactory.AppendRow(myRow, myRowData, TableResponses)
                            Next
                        Next


                        '########################################################
                        'LogVariables
                        '########################################################
                        Dim TableLogVariables As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "LogVariables")



                        '########################################################
                        'TechLog
                        '########################################################
                        Dim TableTechLog As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "TechLog")

                        '########################################################
                        'Dim currentUser As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent
                        'Dim currentUserName As String = currentUser.Name.Substring(currentUser.Name.IndexOf("\") + 1)

                        xlsxFactory.SetCellValueString("A", 1, TableTechLog, "Zeitpunkt-Variablen Testcenter", xlsxFactory.CellFormatting.Null, myStyles)
                        xlsxFactory.SetCellValueString("A", 2, TableTechLog, "generiert mit " + My.Application.Info.AssemblyName + " V" +
                                                       My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString + "." +
                                                       My.Application.Info.Version.Build.ToString + " am " + DateTime.Now.ToShortDateString + " " + DateTime.Now.ToShortTimeString +
                                                       " (" + currentUserName + ")", xlsxFactory.CellFormatting.Null, myStyles)

                        myRow = 4

                        xlsxFactory.SetCellValueString("A", myRow, TableTechLog, "ID", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableTechLog, 30)
                        xlsxFactory.SetCellValueString("B", myRow, TableTechLog, "Start at", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", TableTechLog, 20)
                        xlsxFactory.SetCellValueString("C", myRow, TableTechLog, "loadcomplete after", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("C", TableTechLog, 20)
                        xlsxFactory.SetCellValueString("D", myRow, TableTechLog, "loadspeed", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("D", TableTechLog, 20)

                        myColumn = "E"
                        Columns.Clear()

                        For Each s As String In From v As String In AllUnitsWithResponses Order By v Select v
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableTechLog, s + "##staying", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableResponses, 10)
                            Columns.Add(s, myColumn)
                            myColumn = xlsxFactory.GetNextColumn(myColumn)
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableTechLog, s, xlsxFactory.CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableResponses, 10)
                            myColumn = xlsxFactory.GetNextColumn(myColumn)
                        Next

                        xlsxFactory.SetCellValueString("F", myRow, TableTechLog, "unit history", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("F", TableTechLog, 20)


                        progressMax = myTestPersonList.Count
                        progressCount = 1
                        For Each tc As KeyValuePair(Of String, TestPerson) In myTestPersonList
                            If myworker.CancellationPending Then Exit For
                            myworker.ReportProgress(progressCount * 100 / progressMax, "")
                            progressCount += 1

                            myRow += 1
                            Dim myRowData As New List(Of RowData)
                            myRowData.Add(New RowData With {.Column = "A", .Value = tc.Key, .CellType = CellTypes.str})
                            myRowData.Add(New RowData With {.Column = "B", .Value = tc.Value.firstBookletLoadStart, .CellType = CellTypes.int})
                            myRowData.Add(New RowData With {.Column = "C", .Value = tc.Value.loadtime, .CellType = CellTypes.int})
                            myRowData.Add(New RowData With {.Column = "D", .Value = tc.Value.loadspeed(bookletSize).ToString("#.##0,0"), .CellType = CellTypes.dec})
                            myRowData.Add(New RowData With {.Column = "E", .Value = tc.Value.firstunitentertime, .CellType = CellTypes.int})
                            'myRowData.Add(New RowData With {.Column = "F", .Value = tc.Value.unitActivities.getUnitPeriods(tc.Value.firstBookletLoadStart), .CellType = CellTypes.text})
                            xlsxFactory.AppendRow(myRow, myRowData, TableTechLog)
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
