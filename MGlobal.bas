Attribute VB_Name = "MGlobal"
Option Explicit

Public G_strMdbPath As String


Public Sub RstOpen(Source As String, conToUse As ADODB.Connection, rstToOpen As ADODB.Recordset, CursorType As CursorTypeEnum, LockType As LockTypeEnum, Optional lngCacheSize As Long = 1, Optional ByVal MakeOffline As Boolean = False)
    On Error GoTo ERROR_HANDLER_BOOKMARK
    
START:
    If Not rstToOpen Is Nothing Then
        If rstToOpen.State = adStateOpen Then
            rstToOpen.Close
        End If
        Set rstToOpen = Nothing
    End If
    Set rstToOpen = New ADODB.Recordset
    If MakeOffline = True Then
        rstToOpen.CursorLocation = adUseClient
    End If
    
    rstToOpen.CacheSize = lngCacheSize
        
    'Debug.Print Source
    
    
    rstToOpen.Open Source, conToUse, CursorType, LockType
    On Error GoTo 0
    
    If MakeOffline = True Then
        Set rstToOpen.ActiveConnection = Nothing
    End If
    
    On Error GoTo 0
    
    Exit Sub
    
ERROR_HANDLER_BOOKMARK:
    Select Case Err.Number
        Case -2147467259
            Err.Clear
            Set rstToOpen = Nothing
            GoTo START
        Case Else
            Err.Raise Err.Number, , Err.Description
    End Select
End Sub

Public Sub RstClose(rstToClose As ADODB.Recordset)

    If Not rstToClose Is Nothing Then
        If rstToClose.State = adStateOpen Then
            rstToClose.Close
        End If
        Set rstToClose = Nothing
    End If

End Sub

Public Sub DisconnectDB(conToDisconnect As ADODB.Connection)
    If Not conToDisconnect Is Nothing Then
        If conToDisconnect.State = adStateOpen Then
            conToDisconnect.Close
        End If
        Set conToDisconnect = Nothing
    End If
End Sub

Public Sub ConnectDB(ADOConnection As ADODB.Connection, DBPath As String, DBName As String)
    
    If Not ADOConnection Is Nothing Then
        If ADOConnection.State = adStateOpen Then
            ADOConnection.Close
        End If
        Set ADOConnection = Nothing
    End If
    Set ADOConnection = New ADODB.Connection

    ADOConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & "\" & DBName & ";Persist Security Info=False;Jet OLEDB:Database Password=wack2"
End Sub
