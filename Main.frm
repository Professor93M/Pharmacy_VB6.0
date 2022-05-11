VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Picture         =   "Main.frx":0000
      TabIndex        =   4
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Treatment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   2
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Drugs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   1
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      TabIndex        =   0
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   4440
      Picture         =   "Main.frx":1A3A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drugs Inventory"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   44.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   10335
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Inventory.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
    Drugs.Show
    Me.Hide
End Sub

Private Sub Command3_Click()
    Treatment.Show
    Me.Hide
End Sub

Private Sub Command4_Click()
    End
End Sub

