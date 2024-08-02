VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form inventoryform 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   14790
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4080
      Top             =   5040
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\el monck\Desktop\PROJECT\INVENTORY.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\el monck\Desktop\PROJECT\INVENTORY.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "inventory"
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
   Begin MSComCtl2.DTPicker DTPicker2 
      DataField       =   "DATE IN INVENTORY"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102432769
      CurrentDate     =   45098
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "EXPIRING DATE"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Format          =   102432769
      CurrentDate     =   45098
   End
   Begin VB.TextBox txtwholesaler 
      DataField       =   "WHOLESALER"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txtdrug 
      DataField       =   "DRUGNAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtdescription 
      DataField       =   "DESCRIPTION"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox txtamnt 
      DataField       =   "AMOUNT"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txteffects 
      DataField       =   "POSSIBLE EFFECTS"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdprev 
      BackColor       =   &H00808000&
      Caption         =   "<<PREV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00808000&
      Caption         =   "NEXT>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H000000FF&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H000000FF&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdreupdate 
      BackColor       =   &H000080FF&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H000000FF&
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Wholesaler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lbldrug 
      BackColor       =   &H00FF8080&
      Caption         =   "DRUG NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblamnt 
      BackColor       =   &H00FF8080&
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lbldescription 
      BackColor       =   &H00FF8080&
      Caption         =   "DESCRIPTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "DRUG INVENTORY FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   1800
      TabIndex        =   14
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label lblexpiring 
      BackColor       =   &H00FF8080&
      Caption         =   "EXPIRING DATE(dd/mm/yy)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lbldate 
      BackColor       =   &H00FF8080&
      Caption         =   "DATE (dd/mm/yy)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbleffects 
      BackColor       =   &H00FF8080&
      Caption         =   "POSSIBLE SIDE EFFECTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "inventoryform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
Adodc1.Recordset.AddNew
txtdrug.SetFocus
MsgBox "RECORD ADDED SUCCESSFULLY"

End Sub

Private Sub cmddelete_Click()
Dim response As String
response = MsgBox("ARE YOU SURE YOU WANT TO DELETE?", vbOKCancel + vbInformation)
If response = vbOK Then
Adodc1.Recordset.Delete
Else
inventoryform.Show
End If
End Sub

Private Sub cmdnext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub cmdprev_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub cmdreupdate_Click()
Adodc1.Recordset.Update
MsgBox "record saved successfully", vbInformation
End Sub

Private Sub Command1_Click()
salessearch.Show
End Sub

Private Sub Command2_Click()
reportform.Show
End Sub

Private Sub Command3_Click()
patientsearch.Show
End Sub
