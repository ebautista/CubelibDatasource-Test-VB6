VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form InsertForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Test"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Insert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   8400
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Result"
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   15255
      Begin GridEX16.GridEX GridEX2 
         Height          =   3135
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5530
         ReadOnly        =   -1  'True
         Options         =   -1
         RecordsetType   =   1
         AllowEdit       =   0   'False
         ColumnCount     =   2
         CardCaption1    =   -1  'True
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Insert"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      Begin GridEX16.GridEX GridEX1 
         Height          =   3135
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5530
         Options         =   -1
         RecordsetType   =   1
         ColumnCount     =   2
         CardCaption1    =   -1  'True
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
      End
   End
End
Attribute VB_Name = "InsertForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTargetTable As String
Private mTargetDB As String

Private rstTemp As ADODB.Recordset
Private conTemp As ADODB.Connection

Public Sub ShowForm(ByRef Parent As Form, ByVal TargetDB As String, ByVal TargetTable As String)
    mTargetDB = TargetDB
    mTargetTable = TargetTable
    Me.Show vbModal, Parent
End Sub

Private Sub Form_Load()
    Me.Caption = "Test Insert in " & mTargetTable
    
    ConnectDB conTemp, G_strMdbPath, mTargetDB
    RstOpen "SELECT TOP 1 * FROM [" & mTargetTable & "]", conTemp, rstTemp, adOpenKeyset, adLockOptimistic, , True
    rstTemp.AddNew
    
    Set GridEX1.ADORecordset = rstTemp
    GridEX1.LoadEntireRecordset
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RstClose rstTemp
    DisconnectDB conTemp
End Sub

Private Sub Insert_Click()
    GridEX1.Update
    Set GridEX2.ADORecordset = GridEX1.ADORecordset
    GridEX2.LoadEntireRecordset
End Sub


