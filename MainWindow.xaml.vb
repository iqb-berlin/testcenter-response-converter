Class MainWindow
    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        MyAddMessageDelegate = AddressOf InternalSetMessage
        Me.Title = My.Application.Info.AssemblyName

    End Sub

    Private Sub BtnChangeCsvSource_Click(sender As Object, e As RoutedEventArgs)
        Dim folderpicker As New System.Windows.Forms.FolderBrowserDialog With {.Description = "Wählen des Quellverzeichnisses für die Csv-Dateien",
                                                        .ShowNewFolderButton = False, .SelectedPath = My.Settings.lastfolder_csvSources}
        If folderpicker.ShowDialog() Then
            Me.TBCsvSource.Text = folderpicker.SelectedPath
            My.Settings.lastfolder_csvSources = folderpicker.SelectedPath
        End If
    End Sub

    Private Sub BtnChangeBookletTxtSource_Click(sender As Object, e As RoutedEventArgs)
        MsgBox("Diese Funktion ist noch nicht verfügbar.", vbOK, "Warte warte, bald")
    End Sub
    Private Sub BtnHelp_Click(sender As Object, e As RoutedEventArgs)
        MsgBox("Diese Funktion ist noch nicht verfügbar.", vbOK, "Warte warte, bald")
    End Sub

    Private Sub BtnEditor_Click() Handles BtnEditor.Click
        Try
            Dim myType As Type = GetType(MainWindow)
            Dim txtFN As String = IO.Path.GetTempPath + IO.Path.DirectorySeparatorChar + myType.AssemblyQualifiedName + Guid.NewGuid.ToString + ".txt"
            IO.File.WriteAllBytes(txtFN, System.Text.Encoding.Unicode.GetBytes(MyTB.Text))

            Dim proc As New Process
            With proc.StartInfo
                .FileName = txtFN
                .WindowStyle = ProcessWindowStyle.Normal
            End With
            proc.Start()

            Me.DialogResult = True
        Catch ex As Exception
            Dim msg As String = ex.Message
            If ex.InnerException IsNot Nothing Then msg += vbNewLine + ex.InnerException.Message
            Debug.Print("Error Übertragen Meldungen in Texteditor", msg)
        End Try
    End Sub

    '##################################################################################
#Region "Asynch sending Text to TextBlock"
    '##################################################################################
    Private Delegate Sub AddMessageDelegate(MsgType As MessageType, Message As String)
    Private MyAddMessageDelegate As AddMessageDelegate = Nothing

    Private Enum MessageType
        ParagraphOnly
        NoBreakBefore
        Info
        Warning
        ErrorMsg
        title
        header
    End Enum

    Private Shared Property MessageTypeList As New Dictionary(Of String, MessageType) From {
                                                                {"e", MessageType.ErrorMsg},
                                                                {"w", MessageType.Warning},
                                                                {"i", MessageType.Info},
                                                                {"t", MessageType.title},
                                                                {"h", MessageType.header}}



    Public Sub SetMessage(Optional Message As String = "")
        Dim myMsgType As MessageType = MessageType.Info
        If String.IsNullOrEmpty(Message) Then
            myMsgType = MessageType.ParagraphOnly
        Else
            If Message.Length > 3 Then
                If Message.Substring(1, 1) = ":" Then
                    If MessageTypeList.ContainsKey(Message.Substring(0, 1).ToLower) Then
                        myMsgType = MessageTypeList.Item(Message.Substring(0, 1).ToLower)
                        If Message.Substring(2, 1) = " " Then
                            Message = Message.Substring(3)
                        Else
                            Message = Message.Substring(2)
                        End If
                    End If
                End If
            Else
                myMsgType = MessageType.NoBreakBefore
            End If
        End If

        If MyAddMessageDelegate Is Nothing Then
            InternalSetMessage(myMsgType, Message)
        Else
            Me.MyTB.Dispatcher.Invoke(MyAddMessageDelegate, myMsgType, Message)
        End If
    End Sub

    Private Sub InternalSetMessage(MsgType As MessageType, Message As String)
        MyTB.Text = Message
        Select Case MsgType
            Case MessageType.ParagraphOnly : MyTB.Foreground = Brushes.Black
            Case MessageType.ErrorMsg : MyTB.Foreground = Brushes.Crimson
            Case MessageType.Warning : MyTB.Foreground = Brushes.Orange
            Case MessageType.header : MyTB.Foreground = Brushes.LightSteelBlue
            Case MessageType.title : MyTB.Foreground = Brushes.Khaki
            Case Else : MyTB.Foreground = Brushes.MidnightBlue
        End Select
    End Sub
#End Region

    '##################################################################################
#Region "BackgroundWorkers"
    '##################################################################################
    Private WithEvents ConvertResponses_bw As ComponentModel.BackgroundWorker = Nothing
    Private WithEvents ConvertLogs_bw As ComponentModel.BackgroundWorker = Nothing
    Private Sub BtnCancel_Click() Handles BtnCancel.Click
        If ConvertResponses_bw IsNot Nothing AndAlso ConvertResponses_bw.IsBusy Then
            ConvertResponses_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        End If
        If ConvertLogs_bw IsNot Nothing AndAlso ConvertLogs_bw.IsBusy Then
            ConvertLogs_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        End If
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles ConvertResponses_bw.ProgressChanged, ConvertLogs_bw.ProgressChanged
        If Not String.IsNullOrEmpty(e.UserState) Then Me.SetMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles ConvertResponses_bw.RunWorkerCompleted, ConvertLogs_bw.RunWorkerCompleted
        If e.Cancelled Then Me.SetMessage("durch Nutzer abgebrochen.")
        Me.SetMessage("beendet")
        BtnCancel.IsEnabled = False
    End Sub


#End Region
    '##################################################################################
    Private Sub ConvertResponses_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles ConvertResponses_bw.DoWork
        ConvertResponses.Run(sender, My.Settings.lastfolder_csvSources)
    End Sub
    Private Sub BtnResponsesXlsx_Click(sender As Object, e As RoutedEventArgs)
        BtnCancel.IsEnabled = True
        ConvertResponses_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        ConvertResponses_bw.RunWorkerAsync()
    End Sub
    Private Sub ConvertLogs_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles ConvertLogs_bw.DoWork
        ConvertLogs.Run(sender, My.Settings.lastfolder_csvSources)
    End Sub

    Private Sub BtnLogsXlsx_Click(sender As Object, e As RoutedEventArgs)
        BtnCancel.IsEnabled = True
        ConvertLogs_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        ConvertLogs_bw.RunWorkerAsync()
    End Sub

End Class
