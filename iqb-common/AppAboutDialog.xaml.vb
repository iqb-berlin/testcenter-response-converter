Partial Public Class AppAboutDialog
    Private AppName As String
    Public AlertMessageText As String = Nothing

    Private Sub MeLoaded(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles Me.Loaded
        'Hauptinfo aufrufende Assembly
        Dim myMainAssembly As Reflection.Assembly = System.Reflection.Assembly.GetEntryAssembly
        Dim AppData As Reflection.AssemblyName = myMainAssembly.GetName
        AppName = AppData.Name
        Me.ProductName.Text = AppName
        For Each myAttribute As Object In myMainAssembly.GetCustomAttributes(False)
            If TypeOf (myAttribute) Is Reflection.AssemblyCompanyAttribute Then
                Me.CompanyName.Text = CType(myAttribute, Reflection.AssemblyCompanyAttribute).Company
            ElseIf TypeOf (myAttribute) Is Reflection.AssemblyDescriptionAttribute Then
                Me.Description.Text = CType(myAttribute, Reflection.AssemblyDescriptionAttribute).Description
            End If
        Next

        Dim myAppVersion As String = AppData.Version.ToString
        Dim VersionString As String = myAppVersion
        If myAppVersion.Substring(myAppVersion.LastIndexOf(".")) = ".0" Then VersionString = myAppVersion.Substring(0, myAppVersion.LastIndexOf("."))
        If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
            Dim DeplVersion As String = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            If DeplVersion <> myAppVersion Then
                If DeplVersion.Substring(0, DeplVersion.LastIndexOf(".")) = VersionString Then
                    VersionString += " (" + DeplVersion.Substring(DeplVersion.LastIndexOf(".") + 1) + ")"
                Else
                    VersionString += " (" + DeplVersion + ")"
                End If
            End If

            Dim helpsource As String = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.UpdateLocation.AbsoluteUri
            If helpsource.IndexOf("file://") = 0 Then
                Dim fi As New IO.FileInfo(IO.Path.GetDirectoryName(System.Deployment.Application.ApplicationDeployment.CurrentDeployment.UpdateLocation.LocalPath) + "\helpinfo.xml")
                If fi.Exists Then
                    Try
                        Dim XHelpInfo As XDocument = XDocument.Load(fi.FullName)
                        ICHelpers.ItemsSource = XHelpInfo.Root.Elements
                    Catch ex As Exception
                        ICHelpers.ItemsSource = Nothing
                    End Try
                End If
            Else
                Dim xdummy As XElement = <d><h link=<%= IO.Path.GetDirectoryName(helpsource) %>>Infotext zu dieser Anwendung</h></d>
                ICHelpers.ItemsSource = xdummy.Elements
            End If
        End If

        'Liste mit Hilfe-Links
        Me.Version.Text = VersionString
        If ICHelpers.ItemsSource Is Nothing Then Me.DPHelpers.Visibility = Windows.Visibility.Collapsed

        If String.IsNullOrEmpty(AlertMessageText) Then
            Me.TBAlertMessage.Visibility = Windows.Visibility.Collapsed
            'Liste: alle inzwischen geladenen Assemblies
            Dim mySubAppList As New List(Of String)
            For Each mySubAssembly As Reflection.Assembly In AppDomain.CurrentDomain.GetAssemblies
                Dim SubAppData As Reflection.AssemblyName = mySubAssembly.GetName
                mySubAppList.Add(SubAppData.Name + ", " + SubAppData.Version.ToString)
            Next
            mySubAppList.Sort()
            TBSubAssemblies.Text = String.Join(vbNewLine, mySubAppList)
        Else
            TBSubAssemblies.Visibility = Windows.Visibility.Collapsed
            Me.TBAlertMessage.Text = AlertMessageText
        End If
    End Sub

    Private Sub ButtonClose_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub HyperlinkClick(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim link As Hyperlink = sender
        Try
            Process.Start(link.NavigateUri.ToString)
        Catch ex As Exception
            MsgBox("Konnte Link nicht aufrufen: " + ex.Message, MsgBoxStyle.Exclamation, AppName)
        End Try

    End Sub
End Class
