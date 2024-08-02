VERSION 5.00
Begin VB.Form login 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpassword 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox combouser 
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "LOG IN"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "USER"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdlogin_Click()

End Sub
