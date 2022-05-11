VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Treatment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Treatment"
   ClientHeight    =   10140
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   9480
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   9960
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1720
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
      RecordSource    =   "Treatment"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Drugs.frx":0000
      Height          =   2055
      Left            =   120
      TabIndex        =   31
      Top             =   3720
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   3625
      _Version        =   393216
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
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Drugs.frx":0015
      Height          =   2655
      Left            =   120
      TabIndex        =   30
      Top             =   960
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   4683
      _Version        =   393216
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
      ColumnCount     =   8
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
         DataField       =   "patientName"
         Caption         =   "patientName"
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
         DataField       =   "doctorName"
         Caption         =   "doctorName"
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
         DataField       =   "disease"
         Caption         =   "disease"
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
         DataField       =   "symptoms"
         Caption         =   "symptoms"
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
         DataField       =   "visiteDate"
         Caption         =   "visiteDate"
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
      BeginProperty Column07 
         DataField       =   "drug_id"
         Caption         =   "drug_id"
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
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1544.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   2310
      Left            =   5040
      Style           =   1  'Checkbox
      TabIndex        =   21
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox patientName 
      DataField       =   "patientName"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   12135
      Begin VB.TextBox Text1 
         DataField       =   "symptoms"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   7200
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Main"
         Height          =   615
         Left            =   7440
         TabIndex        =   26
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Exit"
         Height          =   615
         Left            =   9720
         TabIndex        =   25
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox doctorName 
         DataField       =   "doctorName"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2640
         TabIndex        =   20
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Height          =   1935
         Left            =   7200
         TabIndex        =   15
         Top             =   1440
         Width           =   4815
         Begin VB.CommandButton Command1 
            Caption         =   "New"
            Height          =   495
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Add"
            Height          =   495
            Left            =   1320
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Delete"
            Height          =   495
            Left            =   3600
            TabIndex        =   23
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Update"
            Height          =   495
            Left            =   2520
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Last >>"
            Height          =   495
            Left            =   3600
            TabIndex        =   19
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "< Previous"
            Height          =   495
            Left            =   2520
            TabIndex        =   18
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Next >"
            Height          =   495
            Left            =   1320
            TabIndex        =   17
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            Caption         =   "<< First"
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search"
         Height          =   1095
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Width           =   4215
         Begin VB.CommandButton Command5 
            Caption         =   "Search"
            Height          =   495
            Left            =   2640
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox srch 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.TextBox visitDate 
         DataField       =   "visiteDate"
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
         Left            =   10080
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Note 
         DataField       =   "note"
         DataSource      =   "Adodc1"
         Height          =   1215
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox disease 
         DataField       =   "disease"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   4920
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Symptoms"
         Height          =   255
         Left            =   7200
         TabIndex        =   28
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Doctor Name"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Drug ID"
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Visit Date"
         Height          =   255
         Left            =   10080
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Note"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Disease"
         Height          =   255
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Treatment Form"
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
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   11775
   End
End
Attribute VB_Name = "Treatment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If (Adodc1.Recordset.Fields("drug_id").Value <> "") Then
        Adodc2.CommandType = adCmdText
        Adodc2.RecordSource = "select * from drugs where id in (" & Adodc1.Recordset.Fields("drug_id").Value & ")"
        Adodc2.Refresh
    End If
End Sub

Private Sub list1_click()
    If (List1.Text <> "") Then
        Adodc2.RecordSource = "select * from drugs where id = " & Mid(List1.Text, 1, 2) & ""
        Adodc2.Refresh
        DataGrid2.Refresh
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Adodc1.Recordset.AddNew
End Sub

Private Sub Command10_Click()
    Main.Show
    Me.Hide
End Sub

Private Sub Command11_Click()
    End
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    For i = 0 To List1.ListCount - 1
        If (List1.Selected(i)) Then
            sdrugs = sdrugs + ListID(List1.List(i), " ") & " ,"
        End If
    Next
    
    Adodc1.Recordset.Fields("patientName") = patientName
    Adodc1.Recordset.Fields("doctorName") = doctorName
    Adodc1.Recordset.Fields("disease") = disease
    Adodc1.Recordset.Fields("visitDate") = visitDate
    Adodc1.Recordset.Fields("note") = Note
    Adodc1.Recordset.Fields("drug_id") = Mid(sdrugs, 1, Len(sdrugs) - 2)
    Adodc1.Recordset.Update
    MsgBox ("Data Added")
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    For i = 0 To List1.ListCount - 1
        If (List1.Selected(i)) Then
            sdrugs = sdrugs + ListID(List1.List(i), " ") & " ,"
        End If
    Next
    Adodc1.Recordset.Update
    Adodc1.Recordset.Fields("patientName") = patientName
    Adodc1.Recordset.Fields("doctorName") = doctorName
    Adodc1.Recordset.Fields("disease") = disease
    Adodc1.Recordset.Fields("visitDate") = visitDate
    Adodc1.Recordset.Fields("note") = Note
    Adodc1.Recordset.Fields("drug_id") = Mid(sdrugs, 1, Len(sdrugs) - 2)
    Adodc1.Recordset.Update
    MsgBox ("Data Updated")
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    If (MsgBox("Delete it?", vbYesNo, "Delete Record") = vbYes) Then
        Adodc1.Recordset.Delete
    End If
End Sub

Private Sub Command5_Click()
    Adodc1.CommandType = adCmdText
    If (srch <> "") Then
        Adodc1.RecordSource = "select * from Treatment where id = " & srch & ""
    Else
        Adodc1.RecordSource = "select * from Treatment"
    End If
    Adodc1.Refresh
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveNext
End Sub

Private Sub Command8_Click()
    On Error Resume Next
    Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command9_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveLast
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from drugs"
    Adodc2.Refresh
    Adodc2.Recordset.MoveFirst
    With Adodc2.Recordset
        Do Until .EOF
            List1.AddItem ![id] & " - " & ![drugname]
            .MoveNext
        Loop
        If (Adodc1.Recordset.EOF) Then
        Else
            List1.Text = Adodc1.Recordset.Fields("Drug_id")
        End If
    End With
    
End Sub

Public Function ListID(st, e) As String
    For i = 1 To Len(st)
        cut = Mid(st, i, 1)
        If (cut = e) Then
            ListID = id
            id = ""
            Exit For
        Else
            id = id & cut
        End If
    Next
End Function
