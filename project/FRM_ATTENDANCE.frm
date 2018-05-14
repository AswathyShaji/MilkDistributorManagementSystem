VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRM_ATTENDANCE 
   Caption         =   "USER ATTENDANCE"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109969409
         CurrentDate     =   42708
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ABSENT"
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
         Left            =   3840
         TabIndex        =   3
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdpresent 
         Caption         =   "PRESNT"
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
         Left            =   2400
         TabIndex        =   2
         Top             =   2640
         Width           =   1215
      End
      Begin VB.ComboBox combouid 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         TabIndex        =   1
         Text            =   ".....select........"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Left            =   720
         TabIndex        =   6
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label4 
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
         Left            =   720
         TabIndex        =   5
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER ATTENDANCE"
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
         TabIndex        =   4
         Top             =   360
         Width           =   4560
      End
   End
End
Attribute VB_Name = "FRM_ATTENDANCE"
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

Private Sub cmdpresent_Click()
If combouid.Text = ".....select........" Or combouid.Text = "" Then
 MsgBox "select user id", vbCritical
 Else
STRSQL = "SELECT * FROM TBL_ATTENDANCE WHERE U_ID='" & combouid.List(combouid.ListIndex) & "' AND A_DATE='" & DTPicker1 & "' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
MsgBox "ALREADY MARKED", vbCritical
Else
subinsertp
MsgBox "success"
End If
End If
End Sub

Private Sub Command1_Click()
If combouid.Text = ".....select........" Or combouid.Text = "" Then
 MsgBox "select user id", vbCritical
 Else
STRSQL = "SELECT * FROM TBL_ATTENDANCE WHERE U_ID='" & combouid.List(combouid.ListIndex) & "' AND A_DATE='" & DTPicker1 & "' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
MsgBox "ALREADY MARKED", vbCritical
Else
subinserta
MsgBox "success"
End If
End If
End Sub

Private Sub Form_Load()
SUBUID
'lbldate.Caption = DateValue(Now)
End Sub

Public Sub subinsertp()
STRSQL = " INSERT INTO TBL_ATTENDANCE (U_ID,A_DATE,A_STATUS) " _
     & " VALUES ('" & combouid.List(combouid.ListIndex) & "','" & DTPicker1 & "'," _
          & " 'p')"
Set RS = adocn.Execute(STRSQL)

End Sub

Public Sub subinserta()
STRSQL = " INSERT INTO TBL_ATTENDANCE (U_ID,A_DATE,A_STATUS) " _
     & " VALUES ('" & combouid.List(combouid.ListIndex) & "','" & DTPicker1 & "'," _
          & " 'a')"
Set RS = adocn.Execute(STRSQL)
End Sub

Private Sub combouid_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub
