VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_USERSALARY 
   Caption         =   "SALARY"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   6480
      TabIndex        =   8
      Top             =   480
      Width           =   4455
      Begin MSFlexGridLib.MSFlexGrid gridsalary 
         Height          =   2415
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4260
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtbp 
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
         Left            =   2280
         TabIndex        =   5
         Top             =   2040
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
         ItemData        =   "FRM_USERSALARY.frx":0000
         Left            =   2280
         List            =   "FRM_USERSALARY.frx":0002
         TabIndex        =   3
         Text            =   "..............select................."
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblbp 
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
         Left            =   3960
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY"
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
         Left            =   600
         TabIndex        =   4
         Top             =   2040
         Width           =   810
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
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY DETAILS"
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
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   3915
      End
   End
End
Attribute VB_Name = "FRM_USERSALARY"
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

Public Sub subinsert()
STRSQL = " INSERT INTO TBL_USERSALARY (U_ID,US_SALARY) " _
     & " VALUES ('" & combouid.List(combouid.ListIndex) & "','" & txtbp.Text & "' )"
Set RS = adocn.Execute(STRSQL)

End Sub

Private Sub cmdadd_Click()
If combouid.Text = "..............select................." Or combouid.Text = "" Then
 MsgBox "select the user id"
 Else
STRSQL = "SELECT * FROM TBL_USERSALARY WHERE U_ID='" & combouid.List(combouid.ListIndex) & "' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
MsgBox "ALREADY EXIST", vbCritical
Else
If fnValidation = True Then
subinsert
MsgBox "success"
subClear
subclearlabel
subAddToGrid
Else
MsgBox "failed", vbCritical
End If
End If
End If
End Sub

Private Sub cmdupdate_Click()
If fnValidation = True Then
subupdate
MsgBox "success"
cmdadd.Enabled = True
cmdupdate.Enabled = False
subAddToGrid
subClear
subclearlabel
Else
MsgBox "failed to update", vbCritical
End If
End Sub

Private Sub Form_Load()
cmdupdate.Enabled = False
SUBUID
subAddToGrid
End Sub

Private Sub subSetgrid()
    gridsalary.Cols = 4
    gridsalary.Rows = 2
    gridsalary.FixedRows = 1
    gridsalary.TextMatrix(0, 1) = "SL No"
    gridsalary.TextMatrix(0, 2) = "USER ID"
    gridsalary.TextMatrix(0, 3) = "SALARY"
    gridsalary.ColWidth(0) = 0
    gridsalary.ColWidth(1) = 750
    gridsalary.ColWidth(2) = 730
    gridsalary.ColWidth(3) = 730
End Sub

Public Sub subAddToGrid()
    gridsalary.Clear
    subSetgrid
    STRSQL = "select * from TBL_USERSALARY"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridsalary.TextMatrix(i, 0) = RS!US_ID
            gridsalary.TextMatrix(i, 1) = SLNO
            gridsalary.TextMatrix(i, 2) = RS!U_ID
            gridsalary.TextMatrix(i, 3) = RS!US_SALARY
            gridsalary.Rows = gridsalary.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridsalary.Rows = gridsalary.Rows - 1
End Sub


Private Sub gridsalary_Click()
    If gridsalary.Rows > 1 Then
        cmdupdate.Enabled = True
        cmdadd.Enabled = False
        STRSQL = "select * from TBL_USERSALARY where US_ID = '" & gridsalary.TextMatrix(gridsalary.RowSel, 0) & "'"
        Set RS = adocn.Execute(STRSQL)
        Dim X As Integer
        For X = 0 To combouid.ListCount - 1
        If combouid.List(X) = RS!U_ID Then
        combouid.ListIndex = X
        Exit For
        End If
        Next X
        txtbp.Text = RS!US_SALARY
    End If
End Sub

Private Sub subupdate()
 STRSQL = " UPDATE TBL_USERSALARY SET US_SALARY= '" & txtbp.Text & "' where U_ID= ' " & combouid.List(combouid.ListIndex) & "'"
 
 Set RS = adocn.Execute(STRSQL)
End Sub


Public Sub subClear()
txtbp.Text = ""
End Sub
Public Function fnValidation()
Dim ok As Boolean
 If (Not IsNumeric(txtbp.Text)) Then
 lblbp.Visible = True
 ok = False
 Else
 lblbp.Visible = False
 ok = True
 End If
 
If ok = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function
Private Sub txtbp_Change()
If Trim(txtbp.Text) = "" Then
    lblbp.Visible = True
    Else
    lblbp.Visible = False
End If
End Sub
Private Sub subclearlabel()
lblbp.Visible = False
End Sub
Private Sub combouid_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub


