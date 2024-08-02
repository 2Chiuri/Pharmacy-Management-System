VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form progressform 
   Caption         =   "Form1"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6960
      Top             =   1680
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1296
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   960
      Width           =   6375
   End
End
Attribute VB_Name = "progressform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Timer1_Timer()
If ProgressBar1 < 99 Then ProgressBar1 = ProgressBar1 + 1 / 8
If ProgressBar1 = 10 Then Label1.Caption = "Welcome to Munyu Healthcare Pharmacy System"
If ProgressBar1 = 30 Then Label1.Caption = "Please wait..."
If ProgressBar1 = 50 Then Label1.Caption = "Validating Database..."
If ProgressBar1 = 80 Then Label1.Caption = "scanning..."
If ProgressBar1 = 90 Then Label1.Caption = "Almost There.Thanks for your patience"
If ProgressBar1 = 96 Then loginform.Show
End Sub
