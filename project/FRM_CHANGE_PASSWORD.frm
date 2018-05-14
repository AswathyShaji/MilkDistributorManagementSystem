VERSION 5.00
Begin VB.Form FRM_CHANGE_PASSWORD 
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtcpassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtname 
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
         Left            =   2760
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblpwd 
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
         Left            =   4440
         TabIndex        =   14
         Top             =   1560
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
         Left            =   4440
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHANGE PASSWORD"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CURRENT PASSWORD"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   5895
      Begin VB.CommandButton cmdcancel 
         Caption         =   "CANCEL"
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
         Left            =   4200
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdchange 
         Caption         =   "CHANGE"
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
         Left            =   2880
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtpassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtnpassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblcpwd 
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
         Left            =   4560
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblnpwd 
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
         Left            =   4560
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIRM PASSWORD"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   2130
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NEW PASSWORD"
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
         TabIndex        =   7
         Top             =   480
         Width           =   2130
      End
   End
End
Attribute VB_Name = "FRM_CHANGE_PASSWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim STRSQL As String

Private Sub cmdcancel_Click()
    subCleardata
    subclearlabel
    txtname.Text = username
End Sub

Private Sub cmdchange_Click()
If fnValidation And fnValidationtxt Then
    subchangepassword
    MsgBox "Password Changed....."
    subCleardata
    txtname.Text = U_USERNAME
    subclearlabel
Else
    MsgBox "Invalid Password or Username.....", vbCritical
End If
End Sub

Public Function fnValidation()
    Dim ok As Boolean
    If Trim(txtcpassword.Text) = "" Or txtcpassword.Text <> pswd Then
        ok = False
    Else
    If Trim(txtnpassword.Text) = "" Then
        ok = False
    Else
    If Trim(txtpassword.Text) = "" Then
        ok = False
    Else
    If txtnpassword.Text <> txtpassword.Text Then
        ok = False
    Else
    ok = True
    End If
    End If
    End If
    End If
    fnValidation = ok
End Function

Public Sub subchangepassword()
    STRSQL = "update TBL_USERINF set U_PASSWORD='" & txtpassword.Text & "' where U_USERNAME='" & txtname.Text & "'"
    Set RS = adocn.Execute(STRSQL)
End Sub

Public Sub subCleardata()
    txtname.Text = ""
    txtcpassword.Text = ""
    txtnpassword.Text = ""
    txtpassword.Text = ""
End Sub


Private Sub Form_Load()
    Me.Left = 4800
    Me.Top = 1400
    txtname.Text = username
End Sub

Public Function fnValidationtxt()
Dim ok1, ok2, ok3, ok4 As Boolean
If Trim(txtname.Text) = "" Then
 lblname.Visible = True
 ok1 = False
 Else
 lblname.Visible = False
 ok1 = True
 End If
 
If Trim(txtcpassword.Text) = "" Then
 lblpwd.Visible = True
 ok2 = False
 Else
 lblpwd.Visible = False
 ok2 = True
  End If
  
If Trim(txtnpassword.Text) = "" Then
 lblnpwd.Visible = True
 ok3 = False
 Else
 lblnpwd.Visible = False
 ok3 = True
 End If
 
 If Trim(txtpassword.Text) = "" Then
 lblcpwd.Visible = True
 ok4 = False
 Else
 lblcpwd.Visible = False
 ok4 = True
 End If
 
If (ok1 And ok2 And ok3 And ok4) = True Then
fnValidationtxt = True
Else
fnValidationtxt = False
End If
End Function
Private Sub txtname_Change()
If Trim(txtname.Text) = "" Then
    lblname.Visible = True
    Else
    lblname.Visible = False
End If
    
End Sub

Private Sub txtcpassword_Change()
If Trim(txtcpassword.Text) = "" Then
    lblpwd.Visible = True
    Else
    lblpwd.Visible = False
End If
End Sub

Private Sub txtnpassword_Change()
If Trim(txtnpassword.Text) = "" Then
    lblnpwd.Visible = True
    Else
    lblnpwd.Visible = False
End If
End Sub

Private Sub txtpassword_Change()
If Trim(txtpassword.Text) = "" Then
    lblcpwd.Visible = True
    Else
    lblcpwd.Visible = False
End If
End Sub

Private Sub subclearlabel()
lblname.Visible = False
lblpwd.Visible = False
lblcpwd.Visible = False
lblnpwd.Visible = False

End Sub
