Imports Newtonsoft.Json


Class UnitEvent
    Public timestamp As Long = 0
    Public key As String = ""
    Public parameter As String = ""
End Class

Class UnitTimeSpan
    Public millisec As Long = 0
    Public count As Integer = 0
End Class

Class UnitActivities
    Inherits Dictionary(Of String, List(Of UnitEvent))


    Public Function getUnitPeriods(starttime As Long) As String
        Dim myreturn As String = ""
        For Each u As KeyValuePair(Of String, List(Of UnitEvent)) In Me
            'find unit enter events
            Dim ts_unittryleave As List(Of Long) = (From ue As UnitEvent In u.Value Order By ue.timestamp Where ue.key = "UNITTRYLEAVE" Select ue.timestamp).ToList()
            Dim lastenter As Long = Long.MaxValue
            For Each ts As Long In (From ue As UnitEvent In u.Value Order By ue.timestamp Descending Where ue.key = "UNITENTER" Select ue.timestamp)
                Dim leave_ts As Long = (From ue As UnitEvent In u.Value Order By ue.timestamp Where ue.key = "UNITTRYLEAVE" And ue.timestamp > ts And ue.timestamp < lastenter Select ue.timestamp).LastOrDefault
                lastenter = ts
                myreturn += "##" + u.Key + ":" + (ts - starttime).ToString + "/" + (leave_ts - ts).ToString
            Next


            'myreturn += String.Join(";", (From ue As UnitEvent In u.Value
            '                              Let stringified As String = ue.timestamp.ToString + ";" + ue.key + ";" + ue.parameter
            '                              Select stringified
            '                          ).ToList())
        Next
        Return myreturn
    End Function
End Class
Class TestPerson
    Public group As String
    Public login As String
    Public code As String
    Public booklet As String
    Public unitActivities As UnitActivities
    Public sysdata As Dictionary(Of String, String)


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
    Public Sub New(g As String, l As String, c As String, b As String)
        group = g
        login = l
        code = c
        booklet = b
        _firstBookletLoadComplete = Long.MaxValue
        _firstBookletLoadStart = Long.MaxValue
        _firstUnitEnter = Long.MaxValue
        sysdata = Nothing
        unitActivities = New UnitActivities
    End Sub
    Public Function loadspeed(bookletsizelist As Dictionary(Of String, Long)) As Double
        Dim myreturn As Double = 0.0
        If bookletsizelist IsNot Nothing AndAlso bookletsizelist.ContainsKey(booklet) Then
            myreturn = bookletsizelist.Item(booklet) / Me.loadtime
        End If
        Return myreturn
    End Function
    Public Sub AddUnitEvent(timestamp As Long, unit As String, event_key As String, event_parameter As String)
        If Not unitActivities.ContainsKey(unit) Then unitActivities.Add(unit, New List(Of UnitEvent))
        unitActivities.Item(unit).Add(New UnitEvent With {.timestamp = timestamp, .key = event_key, .parameter = event_parameter})
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
    Public Sub SetSysdata(g As String, l As String, c As String, b As String, sysdata As Dictionary(Of String, String))
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).sysdata = sysdata
    End Sub
    Public Sub AddUnitEvent(g As String, l As String, c As String, b As String, timestamp As Long, unit As String, event_key As String, event_parameter As String)
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).AddUnitEvent(timestamp, unit, event_key, event_parameter)
        If (event_key = "UNITENTER") Then Me.Item(g + l + c + b).firstUnitEnter = timestamp
    End Sub
End Class