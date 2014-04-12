VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CubelibDatasource Test Module"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstTables 
      Height          =   3570
      Left            =   4080
      TabIndex        =   8
      Top             =   1080
      Width           =   7335
   End
   Begin VB.TextBox txtDBPath 
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   120
      Width           =   10575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Index           =   3
      Left            =   8520
      TabIndex        =   6
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test &Delete"
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   5
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test &Update"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test &Insert"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblPath 
      Caption         =   "DBPath:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTables 
      Caption         =   "Tables"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblDatabases 
      Caption         =   "Databases"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
        Case 0
            InsertForm.ShowForm Me, File1.List(File1.ListIndex), lstTables.List(lstTables.ListIndex)
            
        Case 1
        
        Case 2
        
        Case 3
            Unload Me
              
    End Select
    
End Sub

Private Sub File1_Click()
    PopulateTableList File1.List(File1.ListIndex)
End Sub

Private Sub Form_Load()
    G_strMdbPath = GetSetting("ClearingPoint", "Settings", "MDBPath")
    
    txtDBPath.Text = G_strMdbPath
    
    File1.Path = G_strMdbPath
    File1.Pattern = "*.mdb"
    
    File1.ListIndex = 0
    
    PopulateTableList File1.List(0)
    
    If lstTables.ListCount > 0 Then
        lstTables.ListIndex = 0
    End If
End Sub

Private Sub PopulateTableList(ByVal DBName As String)

    Dim cat As New ADOX.Catalog
    Dim tbl As ADOX.Table
    
    lstTables.Clear
    
    cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & G_strMdbPath & "\" & DBName & ";Jet OLEDB:Database Password=wack2"
    
    For Each tbl In cat.Tables
        If tbl.Type = "TABLE" Then
            lstTables.AddItem tbl.Name
        End If
    Next
End Sub

Private Sub txtDBPath_Change()
    G_strMdbPath = txtDBPath.Text
End Sub
