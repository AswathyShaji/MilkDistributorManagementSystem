VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRM_USERPAYMENT 
   Caption         =   "PAYMENT"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109969409
         CurrentDate     =   42708
      End
      Begin VB.TextBox txtleaves 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3240
         TabIndex        =   2
         Top             =   4800
         Width           =   2295
      End
      Begin VB.TextBox txtano 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3240
         TabIndex        =   15
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txtpno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3240
         TabIndex        =   14
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtbp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   13
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtsalary 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   12
         Top             =   6000
         Width           =   2295
      End
      Begin VB.ComboBox combouid 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FRM_USERPAYMENT.frx":0000
         Left            =   3240
         List            =   "FRM_USERPAYMENT.frx":0002
         TabIndex        =   1
         Text            =   "..............select................."
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton cmdpay 
         Caption         =   "PAY"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   4
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton cmdsalary 
         Caption         =   "CALCULATE SALARY"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   3
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label lblleaves 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty!"
         BeginProperty Font 
            Name            =   "Baskerville Old Face"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   4920
         TabIndex        =   18
         Top             =   4560
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL SALARY"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   11
         Top             =   6120
         Width           =   1560
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BASIC PAY"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   10
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CASUAL LEAVES"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   9
         Top             =   4920
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF ABSENCE"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   8
         Top             =   3960
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF PRESENCE"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   7
         Top             =   3240
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER ID"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAYMENT DETAILS"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   4410
      End
   End
End
Attribute VB_Name = "FRM_USERPAYMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New Recordset
Dim STRSQL As String
Dim autIn As Integer

Private Sub SUBUID()
STRSQL = "SELECT * FROM TBL_USERINF "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combouid.AddItem (RS!U_ID)
RS.MoveNext
Loop
End Sub

Private Sub cmdpay_Click()
If fnValidation Then
subinsert
subclearlabel
MsgBox "success"
End If
End Sub

Private Sub cmdsalary_Click()
subcalculate
End Sub

Private Sub combouid_Click()
fillSALARY
fillPRESENCE
fillABSENT
End Sub

Private Sub Form_Load()
SUBUID
'lbldate.Caption = DateValue(Now)
End Sub


Public Function fillSALARY()
STRSQL = "SELECT * FROM TBL_USERSALARY where U_ID='" & combouid.List(combouid.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
txtbp.Text = RS!US_SALARY
RS.MoveNext
Loop
End Function

Public Function fillPRESENCE()
STRSQL = "SELECT COUNT(A_STATUS) AS VALUE FROM TBL_ATTENDANCE WHERE U_ID='" & combouid.List(combouid.ListIndex) & "' AND MONTH(A_DATE)='" & Month(Now) & "' AND A_STATUS='P'"
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
txtpno.Text = RS!Value
RS.MoveNext
Loop
End Function


Public Function fillABSENT()
STRSQL = "SELECT COUNT(A_STATUS) AS VALUE1 FROM TBL_ATTENDANCE WHERE U_ID='" & combouid.List(combouid.ListIndex) & "' AND MONTH(A_DATE)='" & Month(Now) & "' AND A_STATUS='a'"
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
txtano.Text = RS!Value1
RS.MoveNext
Loop
End Function

Private Sub subcalculate()
Dim a As Double
Dim b As Double
Dim s As Double
Dim c As Double
a = Val(txtpno.Text)
b = Val(txtleaves.Text)
c = Val(txtbp) / 30
s = a + b
txtsalary = c * s
End Sub


Public Sub subinsert()
STRSQL = " INSERT INTO TBL_USERPAYMENT (U_ID,U_DATE,U_PAYMENT) " _
     & " VALUES ('" & combouid.List(combouid.ListIndex) & "','" & DTPicker1 & "','" & txtsalary.Text & "')"
Set RS = adocn.Execute(STRSQL)

End Sub

Public Function fnValidation()
Dim ok As Boolean
If Trim(txtleaves.Text) = "" Then
 lblleaves.Visible = True
 ok = False
 Else
 lblleaves.Visible = False
 ok = True
 End If
 
If ok = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function
Private Sub txtleaves_Change()
If Trim(txtleaves.Text) = "" Then
    lblleaves.Visible = True
    Else
    lblleaves.Visible = False
End If
    
End Sub
Private Sub subclearlabel()
lblleaves.Visible = False
End Sub
Private Sub combouid_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub

