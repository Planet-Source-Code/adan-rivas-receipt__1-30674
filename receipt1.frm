VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Receipt - SMB Software    Programmed by: Adan Rivas"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog c 
      Left            =   2040
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Info 
      Height          =   480
      Index           =   1
      Left            =   2400
      TabIndex        =   17
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   6120
      Width           =   4455
   End
   Begin VB.OptionButton Info2 
      BackColor       =   &H80000009&
      Caption         =   "Money Order"
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   15
      Top             =   4320
      Width           =   2175
   End
   Begin VB.OptionButton Info2 
      BackColor       =   &H80000009&
      Caption         =   "Check"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   14
      Top             =   4320
      Width           =   1335
   End
   Begin VB.OptionButton Info2 
      BackColor       =   &H80000009&
      Caption         =   "Cash"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Info 
      Height          =   480
      Index           =   4
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   9615
   End
   Begin VB.TextBox Info 
      Height          =   480
      Index           =   3
      Left            =   8760
      TabIndex        =   4
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Info 
      Height          =   480
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   6135
   End
   Begin VB.TextBox Info 
      Height          =   480
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000009&
      Caption         =   "Dollars"
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   12120
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000009&
      Caption         =   "Type of Payment"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   12120
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label info1 
      BackColor       =   &H80000009&
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   11
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Caption         =   "TO"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label info1 
      BackColor       =   &H80000009&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   12120
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   "For  Rent  of"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "*****"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   12120
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   7680
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Recived  From"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "DATE:"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12120
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save as..."
         Shortcut        =   ^T
      End
      Begin VB.Menu print 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu tools 
      Caption         =   "Tools"
      Begin VB.Menu choose1 
         Caption         =   "From To Date"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer
Dim t1 As Integer
Dim t2 As Integer

Private Sub about_Click()
MsgBox "Receipt" + vbNewLine + vbNewLine + "SMB Software" + vbNewLine + "=================" + vbNewLine + "SMB Software programmer: Adan Rivas" + vbNewLine + "Copyright 2002"

End Sub

Private Sub choose1_Click()
Form2.Show

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Thank you for using Receipt Program from SMB Software"

End

End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status = "The date goes here"

End Sub

Private Sub Text2_Change()
Status = "The amount goes here"

End Sub

Private Sub open_Click()
With c
.Filter = "Receipt Files Only(*.rct) | *.rct"

.ShowOpen
End With
Open c.FileName For Input As #1
Input #1, i1, i2, i3, i4, i5, i6, i7, i8, i9, I10, i11
Close #1

Info(0) = i1
Info(1) = i2
Info(2) = i3
Info(3) = i4
Info(4) = i5
info1(0) = i6
info1(1) = i7
Info2(0).Value = i8
Info2(1).Value = i9
Info2(2).Value = ii10
Text5 = i11
MsgBox "File opened"

End Sub

Private Sub print_Click()
Me.PrintForm
End Sub

Private Sub save_Click()
On Error GoTo 10
Open c.FileName For Output As #1
Write #1, Info(0), Info(1), Info(2), Info(3), Info(4), info1(0), info1(1), Info2(0), Info2(1), Info2(2), Text5
Close #1
MsgBox "File saved"
GoTo 100
10
With c
.Filter = "Receipt Files Only(*.rct) | *.rct"

.ShowSave
End With
Open c.FileName For Output As #1
Write #1, Info(0), Info(1), Info(2), Info(3), Info(4), info1(0), info1(1), Info2(0), Info2(1), Info2(2), Text5
Close #1
MsgBox "File saved"
100

End Sub

Private Sub saveas_Click()
With c
.Filter = "Receipt Files Only(*.rct) | *.rct"

.ShowSave
End With
Open c.FileName For Output As #1
Write #1, Info(0), Info(1), Info(2), Info(3), Info(4), info1(0), info1(1), Info2(0), Info2(1), Info2(2), Text5
Close #1
MsgBox "File saved"

End Sub
