Public Class MessageBoxUserControl
    Private Delegate Sub AddMessageDelegate(MsgType As MessageType, Message As String)
    Private Delegate Sub ClearMessagesDelegate()
    Private MyAddMessageDelegate As AddMessageDelegate = Nothing
    Private MyClearMessagesDelegate As ClearMessagesDelegate = Nothing
    Private myScrollViewer As ScrollViewer = Nothing

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


    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        MyAddMessageDelegate = AddressOf InternalAddMessage
        MyClearMessagesDelegate = AddressOf InternalClearMessages
    End Sub

    Public Sub ClearMessages()
        If MyClearMessagesDelegate Is Nothing Then
            InternalClearMessages()
        Else
            Me.MyFlow.Dispatcher.Invoke(MyClearMessagesDelegate)
        End If
    End Sub

    Private Sub InternalClearMessages()
        Me.MyFlow.Document = New FlowDocument With {.FontSize = 12, .FontFamily = New FontFamily("Arial, Century Gothic"),
                                              .TextAlignment = TextAlignment.Left}
    End Sub

    Public Sub AddMessage(Optional Message As String = "")
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
            InternalAddMessage(myMsgType, Message)
        Else
            Me.MyFlow.Dispatcher.Invoke(MyAddMessageDelegate, myMsgType, Message)
        End If
    End Sub

    Private Sub InternalAddMessage(MsgType As MessageType, Message As String)

        Dim myParagraph As Paragraph
        If MsgType = MessageType.ParagraphOnly Then
            myParagraph = GetNewParagraph()
        Else
            If MsgType = MessageType.header OrElse MsgType = MessageType.title Then
                myParagraph = GetNewParagraph()
            Else
                myParagraph = GetLastParagraph()
            End If

            Dim myRun As New Run(Message)
            If MsgType = MessageType.NoBreakBefore Then
                If myParagraph.Inlines.Count > 0 Then myRun.Foreground = myParagraph.Inlines.LastOrDefault.Foreground
            Else
                myParagraph.Inlines.Add(New LineBreak)
                Select Case MsgType
                    Case MessageType.ErrorMsg : myRun.Foreground = Brushes.Crimson
                    Case MessageType.Warning : myRun.Foreground = Brushes.Orange
                    Case MessageType.header
                        myRun.Background = Brushes.LightSteelBlue
                        myRun.FontSize = 14.0
                    Case MessageType.title
                        myRun.Background = Brushes.Khaki
                        myRun.FontSize = 16.0
                    Case Else : myRun.Foreground = Brushes.MidnightBlue
                End Select
            End If
            myParagraph.Inlines.Add(myRun)

        End If
        If myScrollViewer Is Nothing Then
            Dim firstChild As DependencyObject = VisualTreeHelper.GetChild(Me.MyFlow, 0)
            If firstChild IsNot Nothing Then
                Dim border As Decorator = VisualTreeHelper.GetChild(firstChild, 0)
                If border IsNot Nothing Then myScrollViewer = border.Child
            End If
        End If

        If myScrollViewer IsNot Nothing Then myScrollViewer.ScrollToEnd()
    End Sub

    Private Function GetLastParagraph() As Paragraph
        If Me.MyFlow.Document Is Nothing Then Me.MyFlow.Document = New FlowDocument With {.FontSize = 12, .FontFamily = New FontFamily("Arial, Century Gothic"),
                                                      .TextAlignment = TextAlignment.Left}
        Dim myParagraph As Paragraph
        If Me.MyFlow.Document.Blocks.Count = 0 Then
            myParagraph = New Paragraph
            Me.MyFlow.Document.Blocks.Add(myParagraph)
        Else
            myParagraph = Me.MyFlow.Document.Blocks.LastOrDefault
        End If
        Return myParagraph
    End Function

    Private Function GetNewParagraph() As Paragraph
        If Me.MyFlow.Document Is Nothing Then Me.MyFlow.Document = New FlowDocument With {.FontSize = 12, .FontFamily = New FontFamily("Arial, Century Gothic"),
                                                      .TextAlignment = TextAlignment.Left}
        Dim myParagraph As New Paragraph
        Me.MyFlow.Document.Blocks.Add(myParagraph)
        Return myParagraph
    End Function

    Public ReadOnly Property Text() As String
        Get
            If MyFlow.Document Is Nothing Then
                Return Nothing
            Else
                Dim textRange As New TextRange(MyFlow.Document.ContentStart, MyFlow.Document.ContentEnd)
                Return textRange.Text
            End If
        End Get
    End Property

    Public Sub StartEditorWithText()
        Try
            Dim myType As Type = GetType(MainWindow)
            Dim txtFN As String = IO.Path.GetTempPath + IO.Path.DirectorySeparatorChar + myType.AssemblyQualifiedName + Guid.NewGuid.ToString + ".txt"
            IO.File.WriteAllBytes(txtFN, System.Text.Encoding.Unicode.GetBytes(Me.Text))

            Dim proc As New Process
            With proc.StartInfo
                .FileName = txtFN
                .WindowStyle = ProcessWindowStyle.Normal
            End With
            proc.Start()
        Catch ex As Exception
            Dim msg As String = ex.Message
            If ex.InnerException IsNot Nothing Then msg += vbNewLine + ex.InnerException.Message
            Debug.Print("Error Übertragen Meldungen in Texteditor", msg)
        End Try
    End Sub
End Class