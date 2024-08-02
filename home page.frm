VERSION 5.00
Begin VB.Form homepage 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10275
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleMode       =   0  'User
   ScaleWidth      =   11358.86
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   6180
      Left            =   1200
      Picture         =   "home page.frx":0000
      ScaleHeight     =   7345.225
      ScaleMode       =   0  'User
      ScaleWidth      =   9180
      TabIndex        =   5
      Top             =   120
      Width           =   9240
   End
   Begin VB.CommandButton Command5 
      Caption         =   "REPORTS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      MaskColor       =   &H00404040&
      TabIndex        =   4
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PATIENT SEARCH"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      MaskColor       =   &H00404040&
      TabIndex        =   3
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SALES SEARCH"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      MaskColor       =   &H00404040&
      TabIndex        =   2
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "SALES"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      MaskColor       =   &H008080FF&
      TabIndex        =   1
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "INVENTORY"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      MaskColor       =   &H008080FF&
      TabIndex        =   0
      Top             =   7440
      Width           =   1815
   End
End
Attribute VB_Name = "homepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
inventoryform.Show
End Sub

Private Sub Command2_Click()
salesform.Show
End Sub

Private Sub Command3_Click()
salessearch.Show
End Sub

Private Sub Command4_Click()
patientsearch.Show
End Sub

Private Sub Command5_Click()
reportform.Show
End Sub

