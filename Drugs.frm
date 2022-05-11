VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Drugs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Drugs"
   ClientHeight    =   9390
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   6240
      TabIndex        =   24
      Top             =   6600
      Width           =   5535
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton ADD 
         Caption         =   "Add"
         Height          =   495
         Left            =   1560
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Height          =   495
         Left            =   2880
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   495
         Left            =   4200
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<< First"
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Next >"
         Height          =   495
         Left            =   1560
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "< Previous"
         Height          =   495
         Left            =   2880
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Last >>"
         Height          =   495
         Left            =   4200
         TabIndex        =   25
         Top             =   840
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1455
      Left            =   7680
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\J1\Desktop\Drugs Inventory\Drugs Inventory.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\J1\Desktop\Drugs Inventory\Drugs Inventory.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Drugs"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox planetype 
      DataField       =   "note"
      DataSource      =   "Adodc1"
      Height          =   1095
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   6960
      Width           =   5535
   End
   Begin VB.TextBox seatno 
      DataField       =   "qty"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9120
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox tripno 
      DataField       =   "company"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox tripcount 
      DataField       =   "type"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox dest 
      DataField       =   "Drugname"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Treatment Data"
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   11775
      Begin VB.TextBox Text4 
         DataField       =   "effects"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3240
         TabIndex        =   34
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         DataField       =   "price"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         DataField       =   "prodDate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   6120
         TabIndex        =   20
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         DataField       =   "validDate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   9000
         TabIndex        =   19
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Exit"
         Height          =   615
         Left            =   9120
         TabIndex        =   15
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Main"
         Height          =   615
         Left            =   6360
         TabIndex        =   14
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search"
         Height          =   1095
         Left            =   360
         TabIndex        =   13
         Top             =   3960
         Width           =   5535
         Begin VB.TextBox srch 
            Height          =   495
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   2775
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Search"
            Height          =   495
            Left            =   3840
            TabIndex        =   22
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Effects"
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Valid Date"
         Height          =   255
         Left            =   9000
         TabIndex        =   18
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Price"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Production Date"
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Note"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   9000
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Company"
         Height          =   255
         Left            =   6120
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Drug Name"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Treatment.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Drugname"
         Caption         =   "Drugname"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "type"
         Caption         =   "type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "company"
         Caption         =   "company"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "qty"
         Caption         =   "qty"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "price"
         Caption         =   "price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "effects"
         Caption         =   "effects"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "prodDate"
         Caption         =   "prodDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "validDate"
         Caption         =   "validDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "note"
         Caption         =   "note"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1679.811
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drugs Form"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   11775
   End
End
Attribute VB_Name = "Drugs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADD_Click()
    On Error Resume Next
    Adodc1.Recordset.Update
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Adodc1.Recordset.Update
End Sub

Private Sub Command10_Click()
    End
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Adodc1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    If (MsgBox("Delete it?", vbYesNo, "Delete Record") = vbYes) Then
        Adodc1.Recordset.Delete
    End If
End Sub

Private Sub Command4_Click()
    Adodc1.CommandType = adCmdText
    sql = "select * from drugs where id = " & srch & ""
    If (srch <> "") Then
        Adodc1.RecordSource = sql
    Else
        Adodc1.RecordSource = "select * from drugs"
    End If
    Adodc1.Refresh
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveNext
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command8_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveLast
End Sub

Private Sub Command9_Click()
    Main.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    Adodc1.Visible = False
End Sub
