Class ResponseEntry
    Public group As String
    Public login As String
    Public code As String
    Public booklet As String
    Public unit As String
    Public data As Dictionary(Of String, String)
    Public responseTimestamp As String
    Public ReadOnly Property Key As String
        Get
            Return group + login + code
        End Get
    End Property

    Public Shared Function getResponseEntriesFromLine(Line As String, Optional errorLocation As String = "") As List(Of ResponseEntry)
        Dim myreturn As New List(Of ResponseEntry)

        Dim response As String = ""
        Dim responseType As String = ""

        Dim localgroup As String = ""
        Dim locallogin As String = ""
        Dim localcode As String = ""
        Dim localbooklet As String = ""
        Dim localunit As String = ""
        Dim localresponseTimestamp As String = ""

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
                                localgroup = tmpStr
                            Case 1
                                locallogin = tmpStr
                            Case 2
                                localcode = tmpStr
                            Case 3
                                localbooklet = tmpStr
                            Case 4
                                localunit = tmpStr
                            Case 5
                                response = tmpStr
                            Case 7
                                responseType = tmpStr
                            Case 8
                                localresponseTimestamp = tmpStr
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
                Dim localdata As New Dictionary(Of String, String)
                Try
                    localdata = Newtonsoft.Json.JsonConvert.DeserializeObject(tmpResponse, GetType(Dictionary(Of String, String)))
                Catch ex As Exception
                    localdata.Add("ConverterError", "parsing " + responseType + "failed: " + ex.Message)
                    If Not String.IsNullOrEmpty(errorLocation) Then localdata.Add("ErrorLocation", errorLocation)
                    Debug.Print("parseError " + ex.Message + " @ " + errorLocation)
                    Debug.Print(tmpResponse)
                End Try
                myreturn.Add(New ResponseEntry With {
                             .booklet = localbooklet,
                             .code = localcode,
                             .data = localdata,
                             .group = localgroup,
                             .login = locallogin,
                             .responseTimestamp = localresponseTimestamp,
                             .unit = localunit})

                '################## IQBSurveysV1 ################################
            ElseIf responseType = "IQBSurveysV1" Then
                Dim tmpResponse As String = response.Replace("""""", """")
                tmpResponse = tmpResponse.Replace("\\", "\")
                Dim allResponses As String() = Nothing
                Dim localdata As New Dictionary(Of String, String)
                Try
                    allResponses = Newtonsoft.Json.JsonConvert.DeserializeObject(tmpResponse, GetType(String()))
                Catch ex As Exception
                    localdata.Add("ConverterError", "parsing " + responseType + "failed: " + ex.Message)
                    If Not String.IsNullOrEmpty(errorLocation) Then localdata.Add("ErrorLocation", errorLocation)
                    Debug.Print("parseError " + ex.Message + " @ " + errorLocation)
                    Debug.Print(tmpResponse)
                End Try
                If allResponses Is Nothing Then
                    myreturn.Add(New ResponseEntry With {
                             .booklet = localbooklet,
                             .code = localcode,
                             .data = localdata,
                             .group = localgroup,
                             .login = locallogin,
                             .responseTimestamp = localresponseTimestamp,
                             .unit = localunit})
                Else
                    Dim testeeData As New Dictionary(Of Integer, Dictionary(Of String, String))
                    For Each s As String In allResponses
                        Dim sSplits As String() = s.Split({"::"}, StringSplitOptions.RemoveEmptyEntries)
                        If sSplits.Count = 2 Then
                            'find out person
                            Dim pIndex As Integer = 0
                            Dim pPos As Integer = sSplits(0).LastIndexOf("_")
                            If pPos > 1 AndAlso Integer.TryParse(sSplits(0).Substring(pPos + 1), pIndex) Then
                                If pIndex > 0 Then
                                    If Not testeeData.ContainsKey(pIndex) Then testeeData.Add(pIndex, New Dictionary(Of String, String))
                                    Dim varname As String = sSplits(0).Substring(0, pPos)
                                    If Not testeeData.Item(pIndex).ContainsKey(varname) Then testeeData.Item(pIndex).Add(varname, sSplits(1))
                                End If
                            End If
                            If pIndex <= 0 Then localdata.Add(sSplits(0), sSplits(1))
                        End If
                    Next
                    If testeeData.Count > 0 Then
                        For Each td As KeyValuePair(Of Integer, Dictionary(Of String, String)) In testeeData
                            myreturn.Add(New ResponseEntry With {
                             .booklet = localbooklet,
                             .code = IIf(String.IsNullOrEmpty(localcode), td.Key.ToString, localcode + "_" + td.Key.ToString),
                             .data = td.Value,
                             .group = localgroup,
                             .login = locallogin,
                             .responseTimestamp = localresponseTimestamp,
                             .unit = localunit})
                        Next
                    End If
                    If localdata.Count > 0 Then
                        myreturn.Add(New ResponseEntry With {
                                     .booklet = localbooklet,
                                     .code = localcode,
                                     .data = localdata,
                                     .group = localgroup,
                                     .login = locallogin,
                                     .responseTimestamp = localresponseTimestamp,
                                     .unit = localunit})
                    End If
                End If

            Else
                '##################
                Debug.Print("buggy responseType for " + response)
                Dim localdata As New Dictionary(Of String, String)
                If String.IsNullOrEmpty(responseType) Then
                    localdata.Add("ConverterError", "responseType not given")
                Else
                    localdata.Add("ConverterError", "unknown responseType " + responseType)
                End If
                If Not String.IsNullOrEmpty(errorLocation) Then localdata.Add("ErrorLocation", errorLocation)
                myreturn.Add(New ResponseEntry With {
                             .booklet = localbooklet,
                             .code = localcode,
                             .data = localdata,
                             .group = localgroup,
                             .login = locallogin,
                             .responseTimestamp = localresponseTimestamp,
                             .unit = localunit})

            End If
        End If

        Return myreturn
    End Function

End Class

