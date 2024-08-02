VERSION 5.00
Begin VB.Form loginform 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdlogin 
      Caption         =   "LOG IN"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox combouser 
      Height          =   315
      ItemData        =   "loginform.frx":0000
      Left            =   2640
      List            =   "loginform.frx":000A
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtpassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "USER"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "loginform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdlogin_Click()
If combouser.Text = "PHARMACIST" And txtpassword.Text = "1234" Then
homepage.Command3.Enabled = False
homepage.Command4.Enabled = False
homepage.Command5.Enabled = False
homepage.Show

ElseIf combouser.Text = "MANAGER" And txtpassword.Text = "4321" Then
homepage.Show
Else
MsgBox "Wrong Password!"
End If

If txtpassword.Text = "" Then
MsgBox "PLEASE INPUT PASSWORD"
End If
txtpassword.Text = ""
End Sub
