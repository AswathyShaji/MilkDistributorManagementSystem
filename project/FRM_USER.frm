VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_USER 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "USER INFORMATION"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16620
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   16620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   7440
      Width           =   6015
      Begin VB.CommandButton cmddelete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   6120
      TabIndex        =   24
      Top             =   0
      Width           =   10095
      Begin MSFlexGridLib.MSFlexGrid griduser 
         Height          =   6735
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   11880
         _Version        =   393216
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW DETAILS"
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
         Left            =   1440
         TabIndex        =   29
         Top             =   360
         Width           =   3300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtaddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtph 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   3
         Top             =   4080
         Width           =   3135
      End
      Begin VB.TextBox txtemail 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   4920
         Width           =   3135
      End
      Begin VB.TextBox txtusername 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   5
         Top             =   5760
         Width           =   3135
      End
      Begin VB.TextBox txtpassword 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   6600
         Width           =   3135
      End
      Begin VB.Label lblid 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   28
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label lblpassword 
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
         Left            =   4800
         TabIndex        =   23
         Top             =   6360
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblusername 
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
         Left            =   4800
         TabIndex        =   22
         Top             =   5520
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblemail 
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
         Left            =   3720
         TabIndex        =   21
         Top             =   4560
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblph 
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
         Left            =   3720
         TabIndex        =   20
         Top             =   3840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lbladdress 
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
         Left            =   4800
         TabIndex        =   19
         Top             =   2400
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblname 
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
         Left            =   4800
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRATION"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   3405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
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
         Left            =   240
         TabIndex        =   15
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE NUMBER"
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
         Left            =   240
         TabIndex        =   14
         Top             =   4200
         Width           =   1560
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL ID"
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
         Left            =   240
         TabIndex        =   13
         Top             =   5040
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME"
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
         Left            =   240
         TabIndex        =   12
         Top             =   5880
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
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
         Left            =   240
         TabIndex        =   11
         Top             =   6720
         Width           =   1125
      End
   End
End
Attribute VB_Name = "FRM_USER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim STRSQL As String

Public Sub subinsert()
STRSQL = " INSERT INTO TBL_USERINF (U_NAME,U_ADDRESS,U_PH,U_EMAIL,U_USERNAME,U_PASSWORD,U_STATUS) " _
      & " VALUES ('" & txtname.Text & "' , '" & txtaddress.Text & "' , '" & txtph.Text & "' , " _
      & " '" & txtemail.Text & "' , '" & txtusername.Text & "' , '" & txtpassword.Text & "','1')"

Set RS = adocn.Execute(STRSQL)

End Sub
Public Sub subClear()
txtname.Text = ""
txtaddress.Text = ""
txtph.Text = ""
txtemail.Text = ""
txtusername.Text = ""
txtpassword.Text = ""
End Sub
Private Sub cmdadd_Click()
If fnValidation And fnMobileValidation And fnEmailValidation Then
subinsert
subid
MsgBox "Registration Succesfull"
subClear
subclearlabel
subAddToGrid
           Else
            MsgBox "Registration Failed", vbCritical
            End If
           
End Sub

Private Sub cmdclear_Click()
subClear
End Sub
Public Function fnValidation()
Dim ok1, ok2, ok3, ok4, ok5, ok6 As Boolean
If Trim(txtname.Text) = "" Then
 lblname.Visible = True
 ok1 = False
 Else
 lblname.Visible = False
 ok1 = True
 End If
 
If Trim(txtaddress.Text) = "" Then
 lbladdress.Visible = True
 ok2 = False
 Else
 lbladdress.Visible = False
 ok2 = True
  End If
  
If Trim(txtph.Text) = "" Then
 lblph.Visible = True
 ok3 = False
 Else
 lblph.Visible = False
 ok3 = True
 End If
 
If Trim(txtemail.Text) = "" Then
 lblemail.Visible = True
 ok4 = False
 Else
 lblemail.Visible = False
 ok4 = True
 End If
 
If Trim(txtusername.Text) = "" Then
lblusername.Visible = True
 ok5 = False
 Else
 lblusername.Visible = False
 ok5 = True
 End If
 
If Trim(txtpassword.Text) = "" Then
lblpassword.Visible = True
 ok6 = False
 Else
 lblpassword.Visible = False
 ok6 = True
 End If
If (ok1 And ok2 And ok3 And ok4 And ok5 And ok6) = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function



Private Sub cmddelete_Click()
subdelete
MsgBox "deleted"
subClear
subclearlabel
cmddelete.Enabled = False
cmdupdate.Enabled = False
cmdadd.Enabled = True
subAddToGrid

End Sub

Private Sub cmdupdate_Click()
If fnValidation And fnMobileValidation And fnEmailValidation Then
subupdate
subid
MsgBox "Updation Succesfull"
subClear
subclearlabel
cmdupdate.Enabled = False
cmddelete.Enabled = False
cmdadd.Enabled = True
           Else
            MsgBox "Updation Failed", vbCritical
            End If
subAddToGrid
End Sub

Private Sub Form_Load()
subAddToGrid
subid
cmdupdate.Enabled = False
cmddelete.Enabled = False
End Sub


Private Sub griduser_Click()
    If griduser.Rows > 1 Then
        cmdupdate.Enabled = True
        cmddelete.Enabled = True
        cmdadd.Enabled = False
        STRSQL = "select * from TBL_USERINF where U_ID = '" & griduser.TextMatrix(griduser.RowSel, 0) & "'"
        Set RS = adocn.Execute(STRSQL)
        lblid.Caption = RS!U_ID
        txtname.Text = RS!U_NAME
        txtaddress.Text = RS!U_ADDRESS
        txtph.Text = RS!U_PH
        txtemail.Text = RS!U_EMAIL
        txtusername.Text = RS!U_USERNAME
        txtpassword.Text = RS!U_PASSWORD
    End If
End Sub


Private Sub txtname_Change()
If Trim(txtname.Text) = "" Then
    lblname.Visible = True
    Else
    lblname.Visible = False
End If
    
End Sub

Private Sub txtaddress_Change()
If Trim(txtaddress.Text) = "" Then
    lbladdress.Visible = True
    Else
    lbladdress.Visible = False
End If
End Sub

Private Sub txtph_Change()
If Trim(txtph.Text) = "" Then
    lblph.Visible = True
    Else
    lblph.Visible = False
End If
End Sub

Private Sub txtemail_Change()
If Trim(txtemail.Text) = "" Then
    lblemail.Visible = True
    Else
    lblemail.Visible = False
End If
End Sub

Private Sub txtusername_Change()
If Trim(txtusername.Text) = "" Then
    lblusername.Visible = True
    Else
    lblusername.Visible = False
End If
End Sub

Private Sub txtpassword_Change()
If Trim(txtpassword.Text) = "" Then
    lblpassword.Visible = True
    Else
    lblpassword.Visible = False
End If
End Sub
Public Function fnMobileValidation()
        Dim Mobile As String
        Mobile = txtph.Text
        Dim ok As Boolean
        If (IsNumeric(Mobile) And Len(Mobile) = "10") Then
            ok = True
            lblph.Visible = False
        Else
            ok = False
            lblph.Caption = "* Invalid Mobile Number"
            lblph.Visible = True
        End If
        fnMobileValidation = ok
End Function
Public Function fnEmailValidation()
    Dim Email As String
    Dim ok As Boolean
    Email = txtemail.Text
    LCase (Email)
    If (Email Like "*@*.com" Or Email Like "*@*.co.in") Then
        ok = True
        lblemail.Visible = False
    Else
        ok = False
        lblemail.Caption = "* Invalid Email Id"
        lblemail.Visible = True
    End If
     fnEmailValidation = ok
End Function

Private Sub subSetgrid()
    griduser.Cols = 6
    griduser.Rows = 2
    griduser.FixedRows = 1
    griduser.TextMatrix(0, 1) = "SL No"
    griduser.TextMatrix(0, 2) = "Name"
    griduser.TextMatrix(0, 3) = "User Name"
    griduser.TextMatrix(0, 4) = "Phone Number"
    griduser.TextMatrix(0, 5) = "Email id"
    griduser.ColWidth(0) = 0
    griduser.ColWidth(1) = 750
    griduser.ColWidth(2) = 1730
    griduser.ColWidth(3) = 1730
    griduser.ColWidth(4) = 1600
    griduser.ColWidth(5) = 2800
End Sub

Public Sub subAddToGrid()
    griduser.Clear
    subSetgrid
    STRSQL = "select * from TBL_USERINF"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            griduser.TextMatrix(i, 0) = RS!U_ID
            griduser.TextMatrix(i, 1) = SLNO
            griduser.TextMatrix(i, 2) = RS!U_NAME
            griduser.TextMatrix(i, 3) = RS!U_USERNAME
            griduser.TextMatrix(i, 4) = RS!U_PH
            griduser.TextMatrix(i, 5) = RS!U_EMAIL
            griduser.Rows = griduser.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    griduser.Rows = griduser.Rows - 1
End Sub

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_USERINF"
    Set RS = adocn.Execute(STRSQL)
    If RS.BOF = True And RS.EOF = True Then
    lblid.Caption = 1
    Else
        While Not RS.EOF
autIn = RS.Fields(0)
RS.MoveNext
        Wend
        lblid.Caption = autIn + 1
    End If
End Sub

Private Sub subclearlabel()
lblname.Visible = False
lbladdress.Visible = False
lblph.Visible = False
lblemail.Visible = False
lblusername.Visible = False
lblpassword.Visible = False
End Sub

Private Sub subupdate()
 STRSQL = " UPDATE TBL_USERINF SET U_NAME= '" & txtname.Text & "'," _
 & " U_ADDRESS='" & txtaddress.Text & "',U_PH= '" & txtph.Text & "', " _
 & " U_EMAIL= '" & txtemail.Text & "',U_USERNAME= '" & txtusername.Text & "'," _
 & " U_PASSWORD= '" & txtpassword.Text & "' where U_ID=' " & lblid.Caption & "'"
 
 Set RS = adocn.Execute(STRSQL)
End Sub

Private Sub subdelete()
STRSQL = " DELETE FROM TBL_USERINF WHERE U_ID='" & lblid.Caption & "' "
Set RS = adocn.Execute(STRSQL)
End Sub
