VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form receiptform 
   BackColor       =   &H008080FF&
   Caption         =   "Form2"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form2"
   ScaleHeight     =   7740
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   4320
      Top             =   5160
      Width           =   3015
      _ExtentX        =   5318
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\el monck\Desktop\PROJECT\INVENTORY.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\el monck\Desktop\PROJECT\INVENTORY.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "RECEIPT"
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
   Begin VB.TextBox txtid 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2280
      TabIndex        =   25
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   ">>"
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "<<"
      Height          =   495
      Left            =   5400
      TabIndex        =   15
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PRINT RECEIPT"
      Height          =   375
      Left            =   6960
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtcashier 
      DataField       =   "SERVEDBY"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtdate 
      DataField       =   "DATEISSUED"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtbalance 
      DataField       =   "BALANCE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txttotal 
      DataField       =   "TOTAL"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtmethod 
      DataField       =   "PAYMENTMETHOD"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtname 
      DataField       =   "CUSTOMERNAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "ID"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   480
      Width           =   735
   End
   Begin VB.Label drug7 
      BackColor       =   &H8000000C&
      DataField       =   "DRUG7"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label drug6 
      BackColor       =   &H8000000D&
      DataField       =   "DRUG6"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2520
      TabIndex        =   22
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label drug5 
      BackColor       =   &H8000000C&
      DataField       =   "DRUG5"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label drug4 
      BackColor       =   &H8000000D&
      DataField       =   "DRUG4"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label drug3 
      BackColor       =   &H8000000A&
      DataField       =   "DRUG3"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label drug2 
      BackColor       =   &H8000000D&
      DataField       =   "DRUG2"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2520
      TabIndex        =   18
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label drug1 
      BackColor       =   &H8000000A&
      DataField       =   "DRUG1"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "SERVED BY:"
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "DATE"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "BALANCE"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "TOTAL"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "PAYMENT METHOD"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "DRUGS"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "NAME:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "receiptform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub Command1_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Command2_Click()
DataReport1.Show
End Sub


