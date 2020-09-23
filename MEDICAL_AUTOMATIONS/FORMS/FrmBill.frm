VERSION 5.00
Begin VB.Form FrmBill 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Patient Billing"
   ClientHeight    =   8490
   ClientLeft      =   825
   ClientTop       =   1350
   ClientWidth     =   11880
   Icon            =   "FrmBill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\6TH SEM B.C.A PROJECTS\V.B PROJECT\MEDICAL_AUTOMATIONS\DATABASE\HOSPITAL.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Patient"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      Picture         =   "FrmBill.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      Picture         =   "FrmBill.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton CmdCalculator 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Picture         =   "FrmBill.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton CmdFind 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      Picture         =   "FrmBill.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Picture         =   "FrmBill.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text9 
      DataField       =   "Total"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   6720
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      DataField       =   "BedChar"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   5880
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      DataField       =   "LabChar"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   5040
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      DataField       =   "DrugChar"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   14
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      DataField       =   "ConsChar"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      DataField       =   "RegChar"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      DataField       =   "Pat_Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   10
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "PatientId"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "St. Theresa Medical"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   7560
      TabIndex        =   29
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   11880
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   28
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   27
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   26
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   25
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   24
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Private Limited"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8280
      TabIndex        =   18
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bed Charge :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lab Charge :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Drug Charge :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultation Charge :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Charge :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "FrmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCalculator_Click()
FrmCal.Show
End Sub

Private Sub cmdCancel_Click()
Data1.Recordset.Edit
Data1.Recordset.CancelUpdate
CmdSave.Enabled = False
End Sub

Private Sub CmdExit_Click()
Me.Hide
Unload FrmBill
frmmain.Show
'DataAdmn.Recordset.Edit
'frmmain.StatusBar.SimpleText = "Click Menu Item to Begin..."
End Sub

Private Sub CmdFind_Click()
Data1.Recordset.Edit
Data1.Recordset.MoveFirst
Const conMsg = "No matching record was found."
Const conBtn = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
Dim strFindString As String, intUserResponse As String
strFindString = InputBox("Enter patient registration number.", "Find Record")
If strFindString <> "" Then
    Data1.Recordset.FindFirst "PatientId= '" & strFindString & "'"
    If Data1.Recordset.NoMatch = True Then
        intUserResponse = MsgBox(conMsg, conBtn, "Ramkrishna Hospital Pvt Ltd.")
    End If
End If
Text9.SetFocus
CmdFind.SetFocus
End Sub

Private Sub CmdSave_Click()
Text9.SetFocus
Data1.Recordset.Edit
CmdSave.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub Command1_Click()
Data1.Recordset.Edit
CmdSave.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub Form_Paint()
Text2.Text = Date
CmdSave.Enabled = False
cmdCancel.Enabled = False
Text9.SetFocus
End Sub

Private Sub Text1_Change()
CmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Text2_Change()
CmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Text3_Change()
CmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Text4_Change()
CmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Text5_Change()
CmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Text6_Change()
CmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Text7_Change()
CmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Text8_Change()
CmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub Text9_GotFocus()
Text9.Text = Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
End Sub
