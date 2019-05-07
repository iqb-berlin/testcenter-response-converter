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
        Dim myDlg As New AppAboutDialog With {.Title = "Über " + My.Application.Info.AssemblyName, .Owner = Me}
        myDlg.ShowDialog()
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
                                        If parameter.IndexOf("""") = 0 AndAlso parameter.LastIndexOf("""") = parameter.Length - 1 Then parameter = parameter.Substring(1, parameter.Length - 2)
                                        key = key.Substring(0, key.IndexOf(": "))
                                    End If

                                    Select Case key
                                        Case "BOOKLETLOADSTART"
                                            myTestPersonList.SetFirstBookletLoadStart(group, login, code, booklet, timestampInt)


                                            Dim parameterClean As String = parameter.Replace("""""", """")
                                            parameterClean = parameterClean.Replace("\\", "\")
                                            Dim sysdata As Dictionary(Of String, String) = Nothing
                                            Try
                                                sysdata = JsonConvert.DeserializeObject(parameterClean, GetType(Dictionary(Of String, String)))
                                            Catch ex As Exception
                                                sysdata = Nothing
                                                Debug.Print("sysdata json convert failed: " + ex.Message)
                                            End Try
                                            myTestPersonList.SetSysdata(group, login, code, booklet, sysdata)
                                            myTestPersonList.AddLogEvent(group, login, code, booklet, timestampInt, "#BOOKLET#", key, parameter)
                                        Case "BOOKLETLOADCOMPLETE"
                                            myTestPersonList.SetFirstBookletLoadComplete(group, login, code, booklet, timestampInt)
                                            myTestPersonList.AddLogEvent(group, login, code, booklet, timestampInt, "#BOOKLET#", key, parameter)
                                        Case "RESPONSESCOMPLETE", "PRESENTATIONCOMPLETE"
                                            myTestPersonList.AddLogEvent(group, login, code, booklet, timestampInt, unit, key, parameter)
                                        Case "UNITENTER"
                                            myTestPersonList.SetFirstUnitEnter(group, login, code, booklet, timestampInt)
                                            myTestPersonList.AddLogEvent(group, login, code, booklet, timestampInt, unit, key, parameter)
                                        Case Else
                                            myTestPersonList.AddLogEvent(group, login, code, booklet, timestampInt, unit, key, parameter)
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
                                For Each entry As ResponseEntry In ResponseEntry.getResponseEntriesFromLine(line, "file '" + fi.Name + "', line " + lineCount.ToString())
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
                                Next
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
                        myworker.ReportProgress(0.0#, "Schreibe Daten")

                        '########################################################

                        Dim myRow As Integer = 1
                        xlsxFactory.SetCellValueString("A", myRow, TableResponses, "ID", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableResponses, 20)
                        xlsxFactory.SetCellValueString("B", myRow, TableResponses, "Group", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", TableResponses, 10)
                        xlsxFactory.SetCellValueString("C", myRow, TableResponses, "Login+Code", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("C", TableResponses, 10)
                        xlsxFactory.SetCellValueString("D", myRow, TableResponses, "Booklet", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("D", TableResponses, 10)
                        Dim myColumn As String = "E"
                        Dim Columns As New Dictionary(Of String, String)

                        Dim progressMax As Integer = AllVariables.Count
                        Dim progressCount As Integer = 1
                        Dim stepMax As Integer = 5
                        Dim stepCount As Integer = 1
                        Dim progressValue As Double = 0.0#

                        For Each s As String In From v As String In AllVariables Order By v Select v
                            progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                            myworker.ReportProgress(progressValue, "")
                            progressCount += 1
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableResponses, s, xlsxFactory.CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableResponses, 10)
                            Columns.Add(s, myColumn)
                            myColumn = xlsxFactory.GetNextColumn(myColumn)
                        Next

                        Dim BookletUnits As New Dictionary(Of String, List(Of String)) 'für unten TechTable

                        progressMax = AllData.Count
                        progressCount = 1
                        stepCount += 1
                        For Each persondata As KeyValuePair(Of String, Dictionary(Of String, List(Of ResponseEntry))) In AllData
                            If myworker.CancellationPending Then Exit For
                            progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                            myworker.ReportProgress(progressValue, "")
                            progressCount += 1
                            For Each bookletdata As KeyValuePair(Of String, List(Of ResponseEntry)) In persondata.Value
                                Dim resp As ResponseEntry = bookletdata.Value.FirstOrDefault
                                If resp IsNot Nothing Then
                                    myRow += 1
                                    Dim myRowData As New List(Of RowData)
                                    myRowData.Add(New RowData With {.Column = "A", .Value = persondata.Key + bookletdata.Key, .CellType = CellTypes.str})
                                    myRowData.Add(New RowData With {.Column = "B", .Value = resp.group, .CellType = CellTypes.str})
                                    myRowData.Add(New RowData With {.Column = "C", .Value = resp.login + resp.code, .CellType = CellTypes.str})
                                    myRowData.Add(New RowData With {.Column = "D", .Value = bookletdata.Key, .CellType = CellTypes.str})
                                    For Each u As ResponseEntry In bookletdata.Value
                                        If Not BookletUnits.ContainsKey(bookletdata.Key) Then BookletUnits.Add(bookletdata.Key, New List(Of String))
                                        If Not BookletUnits.Item(bookletdata.Key).Contains(u.unit) Then BookletUnits.Item(bookletdata.Key).Add(u.unit)
                                        For Each d As KeyValuePair(Of String, String) In u.data
                                            myRowData.Add(New RowData With {.Column = Columns.Item(u.unit + "##" + d.Key), .Value = d.Value, .CellType = CellTypes.str})
                                        Next
                                    Next
                                    xlsxFactory.AppendRow(myRow, myRowData, TableResponses)
                                End If
                            Next
                        Next


                        '########################################################
                        'TimeOnPage
                        '########################################################
                        progressMax = myTestPersonList.Count
                        progressCount = 1
                        stepCount += 1
                        Dim TableTimeOnPage As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "TimeOnPage")
                        myRow = 1
                        xlsxFactory.SetCellValueString("A", myRow, TableTimeOnPage, "ID", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableTimeOnPage, 20)

                        Dim AllTimeVariables As New List(Of String)
                        Dim AllTimeOnPage As New Dictionary(Of String, List(Of TimeOnPage))
                        Dim BookletMaxVisitedPagesCount As New Dictionary(Of String, Integer) 'für unten TechTable
                        Dim TesteeBookletVisitedPagesCount As New Dictionary(Of String, Integer) 'für unten TechTable
                        Dim TesteeBookletRespondedUnitsCount As New Dictionary(Of String, Integer) 'für unten TechTable
                        For Each tc As KeyValuePair(Of String, TestPerson) In myTestPersonList
                            If myworker.CancellationPending Then Exit For
                            progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                            myworker.ReportProgress(progressValue, "")
                            progressCount += 1
                            If Not AllTimeOnPage.ContainsKey(tc.Key) Then
                                Dim myTimeOnPageList As List(Of TimeOnPage) = tc.Value.GetTimeOnPageList(AllUnitsWithResponses)
                                AllTimeOnPage.Add(tc.Key, myTimeOnPageList)

                                TesteeBookletVisitedPagesCount.Add(tc.Key, myTimeOnPageList.Count)
                                If BookletMaxVisitedPagesCount.ContainsKey(tc.Value.booklet) Then
                                    If BookletMaxVisitedPagesCount.Item(tc.Value.booklet) < myTimeOnPageList.Count Then BookletMaxVisitedPagesCount.Item(tc.Value.booklet) = myTimeOnPageList.Count
                                Else
                                    BookletMaxVisitedPagesCount.Add(tc.Value.booklet, myTimeOnPageList.Count)
                                End If

                                TesteeBookletRespondedUnitsCount.Add(tc.Key, tc.Value.GetResponsesCompleteAllUnitCount(AllUnitsWithResponses))

                                For Each p As TimeOnPage In myTimeOnPageList
                                    If Not AllTimeVariables.Contains(p.page) Then AllTimeVariables.Add(p.page)
                                Next
                            End If
                        Next

                        myColumn = "B"
                        Columns.Clear()
                        For Each s As String In From v As String In AllTimeVariables Order By v
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableTimeOnPage, s + "##topTotal", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableTimeOnPage, 10)
                            Columns.Add(s, myColumn)
                            myColumn = xlsxFactory.GetNextColumn(myColumn)
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableTimeOnPage, s + "##topCount", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableTimeOnPage, 10)
                            myColumn = xlsxFactory.GetNextColumn(myColumn)
                        Next

                        progressMax = AllTimeOnPage.Count
                        progressCount = 1
                        stepCount += 1
                        For Each topList As KeyValuePair(Of String, List(Of TimeOnPage)) In From top As KeyValuePair(Of String, List(Of TimeOnPage)) In AllTimeOnPage Order By top.Key
                            If myworker.CancellationPending Then Exit For
                            progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                            myworker.ReportProgress(progressValue, "")
                            progressCount += 1

                            myRow += 1
                            Dim myRowData As New List(Of RowData)
                            myRowData.Add(New RowData With {.Column = "A", .Value = topList.Key, .CellType = CellTypes.str})
                            For Each top As TimeOnPage In topList.Value
                                myRowData.Add(New RowData With {.Column = Columns.Item(top.page), .Value = top.millisec, .CellType = CellTypes.int})
                                myRowData.Add(New RowData With {.Column = xlsxFactory.GetNextColumn(Columns.Item(top.page)), .Value = top.count, .CellType = CellTypes.int})
                            Next
                            xlsxFactory.AppendRow(myRow, myRowData, TableTimeOnPage)
                        Next



                        '########################################################
                        'TechData
                        '########################################################
                        Dim TableTechData As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "TechData")
                        Dim currentUser As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent
                        Dim currentUserName As String = currentUser.Name.Substring(currentUser.Name.IndexOf("\") + 1)

                        xlsxFactory.SetCellValueString("A", 1, TableTechData, "Antworten und Log-Daten IQB-Testcenter", xlsxFactory.CellFormatting.Null, myStyles)
                        xlsxFactory.SetCellValueString("A", 2, TableTechData, "konvertiert mit " + My.Application.Info.AssemblyName + " V" +
                                                       My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString + "." +
                                                       My.Application.Info.Version.Build.ToString + " am " + DateTime.Now.ToShortDateString + " " + DateTime.Now.ToShortTimeString +
                                                       " (" + currentUserName + ")", xlsxFactory.CellFormatting.Null, myStyles)

                        myRow = 4

                        xlsxFactory.SetCellValueString("A", myRow, TableTechData, "ID", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableTechData, 30)
                        xlsxFactory.SetCellValueString("B", myRow, TableTechData, "Start at", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", TableTechData, 20)
                        xlsxFactory.SetCellValueString("C", myRow, TableTechData, "loadcomplete after", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("C", TableTechData, 20)
                        xlsxFactory.SetCellValueString("D", myRow, TableTechData, "loadspeed", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("D", TableTechData, 20)
                        xlsxFactory.SetCellValueString("E", myRow, TableTechData, "firstUnitEnter after", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("E", TableTechData, 20)
                        xlsxFactory.SetCellValueString("F", myRow, TableTechData, "os", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("F", TableTechData, 20)
                        xlsxFactory.SetCellValueString("G", myRow, TableTechData, "browser", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("G", TableTechData, 20)
                        xlsxFactory.SetCellValueString("H", myRow, TableTechData, "screen", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("H", TableTechData, 20)
                        xlsxFactory.SetCellValueString("I", myRow, TableTechData, "pages visited ratio", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("I", TableTechData, 20)
                        xlsxFactory.SetCellValueString("J", myRow, TableTechData, "units fully responded ratio", xlsxFactory.CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("J", TableTechData, 20)


                        progressMax = myTestPersonList.Count
                        progressCount = 1
                        stepCount += 1
                        For Each tc As KeyValuePair(Of String, TestPerson) In myTestPersonList
                            If myworker.CancellationPending Then Exit For
                            progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                            myworker.ReportProgress(progressValue, "")
                            progressCount += 1

                            myRow += 1
                            Dim myRowData As New List(Of RowData)
                            myRowData.Add(New RowData With {.Column = "A", .Value = tc.Key, .CellType = CellTypes.str})
                            myRowData.Add(New RowData With {.Column = "B", .Value = tc.Value.firstBookletLoadStart, .CellType = CellTypes.int})
                            myRowData.Add(New RowData With {.Column = "C", .Value = tc.Value.loadtime, .CellType = CellTypes.int})
                            myRowData.Add(New RowData With {.Column = "D", .Value = tc.Value.loadspeed(bookletSize).ToString(), .CellType = CellTypes.dec})
                            myRowData.Add(New RowData With {.Column = "E", .Value = tc.Value.firstUnitEnter - tc.Value.firstBookletLoadStart, .CellType = CellTypes.int})
                            myRowData.Add(New RowData With {.Column = "F", .Value = tc.Value.os, .CellType = CellTypes.str})
                            myRowData.Add(New RowData With {.Column = "G", .Value = tc.Value.browser, .CellType = CellTypes.str})
                            myRowData.Add(New RowData With {.Column = "H", .Value = tc.Value.screen, .CellType = CellTypes.str})

                            Dim myRatio As Double = 0.0#
                            If TesteeBookletVisitedPagesCount.ContainsKey(tc.Key) AndAlso BookletMaxVisitedPagesCount.ContainsKey(tc.Value.booklet) Then
                                Dim bmvpc As Integer = BookletMaxVisitedPagesCount.Item(tc.Value.booklet)
                                If bmvpc > 0 Then myRatio = TesteeBookletVisitedPagesCount.Item(tc.Key) * 100 / bmvpc
                            End If
                            myRowData.Add(New RowData With {.Column = "I", .Value = myRatio.ToString(), .CellType = CellTypes.dec})

                            myRatio = 0.0#
                            If TesteeBookletRespondedUnitsCount.ContainsKey(tc.Key) AndAlso BookletUnits.ContainsKey(tc.Value.booklet) Then
                                Dim buc As Integer = BookletUnits.Item(tc.Value.booklet).Count
                                If buc > 0 Then myRatio = TesteeBookletRespondedUnitsCount.Item(tc.Key) * 100 / buc
                            End If
                            myRowData.Add(New RowData With {.Column = "J", .Value = myRatio.ToString(), .CellType = CellTypes.dec})

                            xlsxFactory.AppendRow(myRow, myRowData, TableTechData)
                        Next


                    End Using
                    myworker.ReportProgress(100.0#, "Speichern Datei")
                    Try
                        Using fs As New IO.FileStream(targetXlsxFilename, IO.FileMode.Create)
                            MemStream.WriteTo(fs)
                        End Using
                    Catch ex As Exception
                        myworker.ReportProgress(100.0#, "e: Konnte Datei nicht schreiben: " + ex.Message)
                    End Try
                End Using
            End If
        End If
    End Sub

End Class
