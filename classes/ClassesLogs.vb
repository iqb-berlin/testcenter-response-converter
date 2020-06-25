Imports Newtonsoft.Json


Class LogEvent
    Public unit As String = ""
    Public key As String = ""
    Public parameter As String = ""
End Class

Class TimeOnPage
    Public page As String = ""
    Public millisec As Long = 0
    Public count As Integer = 0
End Class

Class Activities
    Inherits Dictionary(Of Long, List(Of LogEvent))


    Public Function getUnitPeriods(starttime As Long) As String
        Dim myreturn As String = ""
        'For Each u As KeyValuePair(Of String, List(Of LogEvent)) In Me
        '    'find unit enter events
        '    Dim ts_unittryleave As List(Of Long) = (From ue As LogEvent In u.Value Order By ue.timestamp Where ue.key = "UNITTRYLEAVE" Select ue.timestamp).ToList()
        '    Dim lastenter As Long = Long.MaxValue
        '    For Each ts As Long In (From ue As LogEvent In u.Value Order By ue.timestamp Descending Where ue.key = "UNITENTER" Select ue.timestamp)
        '        Dim leave_ts As Long = (From ue As LogEvent In u.Value Order By ue.timestamp Where ue.key = "UNITTRYLEAVE" And ue.timestamp > ts And ue.timestamp < lastenter Select ue.timestamp).LastOrDefault
        '        lastenter = ts
        '        myreturn += "##" + u.Key + ":" + (ts - starttime).ToString + "/" + (leave_ts - ts).ToString
        '    Next


        '    'myreturn += String.Join(";", (From ue As UnitEvent In u.Value
        '    '                              Let stringified As String = ue.timestamp.ToString + ";" + ue.key + ";" + ue.parameter
        '    '                              Select stringified
        '    '                          ).ToList())
        'Next
        Return myreturn
    End Function
End Class
Class TestPerson
    Public group As String
    Public login As String
    Public code As String
    Public booklet As String
    Public log As Activities

    Public ReadOnly Property loadtime() As Long
        Get
            Return _firstBookletLoadComplete - _firstBookletLoadStart
        End Get
    End Property
    Public ReadOnly Property firstunitentertime() As Long
        Get
            Return _firstUnitEnter - _firstBookletLoadStart
        End Get
    End Property
    Private _firstBookletLoadStart As Long
    Public Property firstBookletLoadStart() As Long
        Get
            Return _firstBookletLoadStart
        End Get
        Set(ByVal value As Long)
            If value < _firstBookletLoadStart Then _firstBookletLoadStart = value
        End Set
    End Property
    Private _firstBookletLoadComplete As Long
    Public Property firstBookletLoadComplete() As Long
        Get
            Return _firstBookletLoadComplete
        End Get
        Set(ByVal value As Long)
            If value < _firstBookletLoadComplete Then _firstBookletLoadComplete = value
        End Set
    End Property
    Private _firstUnitEnter As Long
    Public Property firstUnitEnter() As Long
        Get
            Return _firstUnitEnter
        End Get
        Set(ByVal value As Long)
            If value < _firstUnitEnter Then _firstUnitEnter = value
        End Set
    End Property

    Private _browser As String
    Public ReadOnly Property browser() As String
        Get
            Return Me._browser
        End Get
    End Property
    Private _os As String
    Public ReadOnly Property os() As String
        Get
            Return Me._os
        End Get
    End Property
    Private _screen As String
    Public ReadOnly Property screen() As String
        Get
            Return Me._screen
        End Get
    End Property

    Public Sub New(g As String, l As String, c As String, b As String)
        group = g
        login = l
        code = c
        booklet = b
        _firstBookletLoadComplete = Long.MaxValue
        _firstBookletLoadStart = Long.MaxValue
        _firstUnitEnter = Long.MaxValue
        log = New Activities
        _browser = "?"
        _os = "?"
        _screen = "?"
    End Sub
    Public Function loadspeed(bookletsizelist As Dictionary(Of String, Long)) As Double
        Dim myreturn As Double = 0.0
        If bookletsizelist IsNot Nothing AndAlso bookletsizelist.ContainsKey(booklet) Then
            myreturn = bookletsizelist.Item(booklet) / Me.loadtime
        End If
        Return myreturn
    End Function
    Public Sub AddLogEvent(timestamp As Long, unit As String, event_key As String, event_parameter As String)
        If Not log.ContainsKey(timestamp) Then log.Add(timestamp, New List(Of LogEvent))
        log.Item(timestamp).Add(New LogEvent With {.unit = unit, .key = event_key, .parameter = event_parameter})
    End Sub
    Public Function GetTimeOnPageList(unitsOnly As List(Of String)) As List(Of TimeOnPage)
        Dim unitPageList As New Dictionary(Of String, TimeOnPage)
        Dim currentPage As String = ""
        Dim pageStart As Long = 0
        For Each logList As KeyValuePair(Of Long, List(Of LogEvent)) In From le As KeyValuePair(Of Long, List(Of LogEvent)) In Me.log Order By le.Key
            For Each logEntry As LogEvent In logList.Value
                If logEntry.key = "PAGENAVIGATIONCOMPLETE" AndAlso unitsOnly.Contains(logEntry.unit) Then
                    currentPage = logEntry.unit + "##" + logEntry.parameter
                    pageStart = logList.Key
                ElseIf Not {"RESPONSESCOMPLETE", "PRESENTATIONCOMPLETE", "UNITTRYLEAVE"}.Contains(logEntry.key) Then
                    If Not String.IsNullOrEmpty(currentPage) AndAlso pageStart > 0 Then
                        If Not unitPageList.ContainsKey(currentPage) Then
                            unitPageList.Add(currentPage, New TimeOnPage With {.page = currentPage, .count = 1, .millisec = logList.Key - pageStart})
                        Else
                            Dim myTimeOnPage As TimeOnPage = unitPageList.Item(currentPage)
                            myTimeOnPage.count += 1
                            myTimeOnPage.millisec += logList.Key - pageStart
                        End If
                        currentPage = ""
                        pageStart = 0
                    End If
                End If
            Next
        Next
        If Not String.IsNullOrEmpty(currentPage) AndAlso pageStart > 0 AndAlso unitPageList.ContainsKey(currentPage) Then
            Dim myTimeOnPage As TimeOnPage = unitPageList.Item(currentPage)
            myTimeOnPage.count += 1
            myTimeOnPage.millisec = 0
        End If

        Return (From top As KeyValuePair(Of String, TimeOnPage) In unitPageList Select top.Value).ToList
    End Function
    Public Function GetResponsesCompleteAllUnitCount(unitsOnly As List(Of String)) As Integer
        Dim unitList As New List(Of String)
        For Each logList As KeyValuePair(Of Long, List(Of LogEvent)) In Me.log
            For Each logEntry As LogEvent In logList.Value
                If logEntry.key = "RESPONSESCOMPLETE" AndAlso logEntry.parameter = "all" AndAlso unitsOnly.Contains(logEntry.unit) AndAlso Not unitList.Contains(logEntry.unit) Then
                    unitList.Add(logEntry.unit)
                End If
            Next
        Next

        Return unitList.Count
    End Function
    Public Sub SetSysdata(sysdata As Dictionary(Of String, String))
        _browser = "?"
        _os = "?"
        _screen = "?"
        If sysdata IsNot Nothing Then
            If sysdata.ContainsKey("browserVersion") AndAlso sysdata.ContainsKey("browserName") Then _browser = sysdata.Item("browserName") + " " + sysdata.Item("browserVersion")
            If sysdata.ContainsKey("osName") Then _os = sysdata.Item("osName")
            If sysdata.ContainsKey("screenSizeWidth") AndAlso sysdata.ContainsKey("screenSizeHeight") Then _screen = sysdata.Item("screenSizeWidth") + " x " + sysdata.Item("screenSizeHeight")
        End If
    End Sub

End Class

Class TestPersonList
    Inherits SortedDictionary(Of String, TestPerson)
    Public Sub SetFirstBookletLoadStart(g As String, l As String, c As String, b As String, value As Long)
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).firstBookletLoadStart = value
    End Sub
    Public Sub SetFirstBookletLoadComplete(g As String, l As String, c As String, b As String, value As Long)
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).firstBookletLoadComplete = value
    End Sub
    Public Sub SetFirstUnitEnter(g As String, l As String, c As String, b As String, value As Long)
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).firstUnitEnter = value
    End Sub
    Public Sub SetSysdata(g As String, l As String, c As String, b As String, sysdata As Dictionary(Of String, String))
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).SetSysdata(sysdata)
    End Sub
    Public Sub AddLogEvent(g As String, l As String, c As String, b As String, timestamp As Long, unit As String, event_key As String, event_parameter As String)
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).AddLogEvent(timestamp, unit, event_key, event_parameter)
    End Sub
End Class