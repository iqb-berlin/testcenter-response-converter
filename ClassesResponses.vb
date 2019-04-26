Class ResponseEntry
    Public group As String
    Public login As String
    Public code As String
    Public booklet As String
    Public unit As String
    Private response As String
    Private responseType As String
    Public data As Dictionary(Of String, String)
    Public responseTimestamp As String
    Public ReadOnly Property Key As String
        Get
            Return group + login + code
        End Get
    End Property
    Public Sub New(line As String, Optional errorLocation As String = "")
        group = ""
        login = ""
        code = ""
        booklet = ""
        unit = ""
        response = ""
        responseType = ""
        responseTimestamp = ""
        data = New Dictionary(Of String, String)

        Dim position As Integer = 0
        Dim semicolonActive As Boolean = True
        Dim tmpStr As String = ""
        For Each c As Char In line
            If c = ";" Then
                If semicolonActive Then
                    If Not String.IsNullOrEmpty(tmpStr) Then
                        If tmpStr.Substring(0, 1) = """" Then
                            If tmpStr.Substring(tmpStr.Length - 1, 1) = """" Then
                                tmpStr = tmpStr.Substring(1, tmpStr.Length - 2)
                            End If
                        End If
                        Select Case position
                            Case 0
                                group = tmpStr
                            Case 1
                                login = tmpStr
                            Case 2
                                code = tmpStr
                            Case 3
                                booklet = tmpStr
                            Case 4
                                unit = tmpStr
                            Case 5
                                response = tmpStr
                            Case 7
                                responseType = tmpStr
                            Case 8
                                responseTimestamp = tmpStr
                        End Select
                        tmpStr = ""
                    End If
                    position += 1
                Else
                    tmpStr += c
                End If
            Else
                tmpStr += c
                If c = """" Then semicolonActive = Not semicolonActive
            End If
        Next
        'ignore laststate

        If Not String.IsNullOrEmpty(response) Then

            '################## VERAOnlineV1 ################################
            If responseType = "VERAOnlineV1" Then
                Dim tmpResponse As String = response.Replace("""""", """")
                tmpResponse = tmpResponse.Replace("\\", "\")
                Try
                    data = Newtonsoft.Json.JsonConvert.DeserializeObject(tmpResponse, GetType(Dictionary(Of String, String)))
                Catch ex As Exception
                    data.Add("ConverterError", "parsing " + responseType + "failed: " + ex.Message)
                    If Not String.IsNullOrEmpty(errorLocation) Then data.Add("ErrorLocation", errorLocation)
                    Debug.Print("parseError " + ex.Message + " @ " + errorLocation)
                    Debug.Print(tmpResponse)
                End Try
            Else
                Debug.Print("buggy responseType for " + response)
                If String.IsNullOrEmpty(responseType) Then
                    data.Add("ConverterError", "responseType not given")
                Else
                    data.Add("ConverterError", "unknown responseType " + responseType)
                End If
                If Not String.IsNullOrEmpty(errorLocation) Then data.Add("ErrorLocation", errorLocation)
            End If
        End If
    End Sub

End Class

