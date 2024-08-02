VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form salesform 
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
   Begin VB.TextBox Text1 
      DataField       =   "dateofissue"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   840
      TabIndex        =   44
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4800
      Top             =   5520
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "SALES"
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
   Begin VB.CommandButton cmdupdate 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   10320
      TabIndex        =   24
      Top             =   3960
      Width           =   975
   End
   Begin VB.ComboBox combocashier 
      DataField       =   "SERVED BY"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":0000
      Left            =   2160
      List            =   "salesform.frx":000A
      TabIndex        =   23
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton TOTAL 
      BackColor       =   &H0080FF80&
      Caption         =   "TOTAL"
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
      Left            =   10680
      TabIndex        =   22
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox combodrug1 
      DataField       =   "DRUG 1"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":002A
      Left            =   1920
      List            =   "salesform.frx":003D
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtcustname 
      DataField       =   "CUSTOMERNAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   840
      Width           =   2415
   End
   Begin VB.ComboBox combodrug2 
      DataField       =   "DRUG 2"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":007D
      Left            =   1920
      List            =   "salesform.frx":0090
      TabIndex        =   19
      Top             =   2280
      Width           =   2055
   End
   Begin VB.ComboBox combodrug3 
      DataField       =   "DRUG 3"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":00D0
      Left            =   1920
      List            =   "salesform.frx":00E3
      TabIndex        =   18
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ComboBox combodrug4 
      DataField       =   "DRUG 4"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":0123
      Left            =   1920
      List            =   "salesform.frx":0136
      TabIndex        =   17
      Top             =   3600
      Width           =   2055
   End
   Begin VB.ComboBox combodrug5 
      DataField       =   "DRUG 5"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":0176
      Left            =   2040
      List            =   "salesform.frx":0189
      TabIndex        =   16
      Top             =   4440
      Width           =   1935
   End
   Begin VB.ComboBox combodrug6 
      DataField       =   "DRUG 6"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":01C9
      Left            =   7200
      List            =   "salesform.frx":01DC
      TabIndex        =   15
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox combodrug7 
      DataField       =   "DRUG 7"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":021C
      Left            =   7320
      List            =   "salesform.frx":022F
      TabIndex        =   14
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox amntdrug1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox amntdrug2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox amntdrug3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox amntdrug4 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   4440
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox amntdrug5 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   4440
      TabIndex        =   9
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox amntdrug6 
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox amntdrug7 
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdsubmit 
      BackColor       =   &H0000FF00&
      Caption         =   "ADD NEW"
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
      Index           =   0
      Left            =   5520
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox combomethod 
      DataField       =   "PAYMENT METHOD"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "salesform.frx":026F
      Left            =   7320
      List            =   "salesform.frx":027C
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtpaid 
      Height          =   405
      Left            =   7200
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtid 
      DataField       =   "IDNUMBER"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NEXT>>"
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
      Left            =   8640
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "<<PREV"
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
      Left            =   8640
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H000000C0&
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
      Height          =   495
      Left            =   7080
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H008080FF&
      Caption         =   "SERVED BY:"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   42
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lbltotal 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "TOTAL"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   10920
      TabIndex        =   41
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "CUSTOMER NAME"
      Height          =   375
      Left            =   240
      TabIndex        =   40
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   "DRUG 1"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "DRUG 2"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      Caption         =   "DRUG 3"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      Caption         =   "DRUG 4"
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H008080FF&
      Caption         =   "DRUG 5"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H008080FF&
      Caption         =   "DRUG 6"
      Height          =   495
      Left            =   5760
      TabIndex        =   34
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H008080FF&
      Caption         =   "DRUG 7"
      Height          =   495
      Left            =   5760
      TabIndex        =   33
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "PRICE"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   4200
      TabIndex        =   32
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "PRICE"
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   9600
      TabIndex        =   31
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H008080FF&
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   5640
      TabIndex        =   30
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H008080FF&
      Caption         =   "BALANCE"
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblbalance 
      DataField       =   "BALANCE"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   7440
      TabIndex        =   28
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H008080FF&
      Caption         =   "AMNT PAID"
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "*****WELCOME  TO MUNYU HEALTHCARE PHARMACY*****"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   2400
      TabIndex        =   26
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label16 
      BackColor       =   &H008080FF&
      Caption         =   "ID No."
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "salesform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddelete_Click()
Dim response As String
response = MsgBox("ARE YOU SURE YOU WANT TO DELETE?", vbOKCancel + vbInformation)
If response = vbOK Then
Adodc1.Recordset.Delete
Else
salesform.Show
End If
End Sub

Private Sub cmdnext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub cmdprev_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
Adodc1.Recordset.MoveLast

End If
End Sub



Private Sub cmdsubmit_Click(Index As Integer)
Adodc1.Recordset.AddNew
Dim currentdate As Date
Dim myrecord As String
currentdate = Now
Text1.Text = Format(currentdate, "mm/dd/yyyy")
End Sub

Private Sub cmdupdate_Click(Index As Integer)
Adodc1.Recordset.Update
MsgBox "record updated successfully"
receiptform.txtdate.Text = Text1.Text
receiptform.txtid.Text = txtid.Text
receiptform.txtname.Text = txtcustname
receiptform.txtmethod.Text = combomethod
receiptform.txttotal.Text = lbltotal
receiptform.txtbalance.Text = lblbalance
receiptform.drug1.Caption = combodrug1
receiptform.drug2.Caption = combodrug2
receiptform.drug3.Caption = combodrug3
receiptform.drug4.Caption = combodrug4
receiptform.drug5.Caption = combodrug5
receiptform.drug6.Caption = combodrug6
receiptform.drug7.Caption = combodrug7
receiptform.txtcashier.Text = combocashier
receiptform.Show
End Sub



Private Sub Command1_Click()
Dim currentQuantity As Integer
Dim drug1 As String
drug1 = combodrug1.Text
Adodc2.RecordSource = "SELECT * FROM Inventory WHERE DRUGNAME = '" & drug1 & "'"
Adodc2.Refresh
    If Not Adodc2.Recordset.EOF Then
currentQuantity = Adodc2.Recordset.Fields("AMOUNT").Value
currentQuantity = currentQuantity - 1
Adodc2.Recordset.Fields("Quantity").Value = currentQuantity
Adodc2.Recordset.Update
MsgBox "DRUGS DEDUCTED SUCCESSFULLY"
End If
End Sub

Private Sub TOTAL_Click()
If txtcustname.Text = "" Then
MsgBox "Please enter the name and ID of the customer"
ElseIf txtid.Text = "" Then
MsgBox "Please enter the name and ID of the customer"
End If

Dim sum As Double
Dim balance As Double
sum = 0
sum = sum + Val(amntdrug1.Text) + Val(amntdrug2.Text) + Val(amntdrug3.Text) + Val(amntdrug4.Text) + Val(amntdrug5.Text) + Val(amntdrug6.Text) + Val(amntdrug7.Text)
lbltotal.Caption = sum
balance = Val(txtpaid.Text) - Val(lbltotal.Caption)
lblbalance.Caption = balance
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
txtid.Locked = True
Else
txtid.Locked = False
End If
If KeyAscii = 8 Then
txtid.Locked = False

End If
End Sub
