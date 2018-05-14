VERSION 5.00
Begin VB.Form FRM_LOGIN 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LOGIN"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Show Password"
      Height          =   270
      Left            =   1800
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtusername 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   0
      Picture         =   "FRM_LOGIN.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7365
   End
End
Attribute VB_Name = "FRM_LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Private Sub sublogin()
    If (txtusername.Text = "admin") And (txtpassword.Text = "admin") Then
        username = "admin"
      MDIForm1.NEW.Visible = True
'        MDIForm1.MCOLLECTION.Visible = False
'        MDIForm1.MTYPE.Visible = False
'        MDIForm1.MSALE.Visible = False
'        MDIForm1.FSALE.Visible = False
'        MDIForm1.MSTOCK.Visible = False
         MDIForm1.Show
        subclear
    Else
        STRSQL = "select * from TBL_USERINF where U_USERNAME='" & txtusername.Text & "' and " _
                & " U_PASSWORD='" & txtpassword.Text & "'"
        Set RS = adocn.Execute(STRSQL)
        If (RS.RecordCount > 0) Then
            username = RS!U_USERNAME
            pswd = RS!U_PASSWORD
            MDIForm1.ADD.Enabled = False
            MDIForm1.PAYMENT.Enabled = False
            MDIForm1.CATEGORY.Enabled = Fals
           ' MDIForm1.UPDATEFEED.Enabled = False
            MDIForm1.FEED_DETALIS.Enabled = False
            MDIForm1.CHART.Enabled = False
            'MDIForm1.SALE_REPORT.Enabled = False
            'MDIForm1.STOCK_REPORT.Enabled = False
           ' MDIForm1.CHANGE_PASSWORD.Enabled = False
            MDIForm1.Show
            subclear
        Else
            MsgBox "Invalid Username and Password", vbOKOnly + vbInformation, "Warning"
        End If
    End If
End Sub
Private Function Fnvalidatiion()
    Dim ok As Boolean
    If (txtusername.Text = "") Then
        ok = False
    ElseIf (txtusername.Text = "") Then
        ok = False
    Else
        ok = True
    End If
    Fnvalidatiion = ok
End Function

Private Sub subclear()
    txtusername.Text = ""
    txtpassword.Text = ""
End Sub

Private Sub Check1_Click()
If Check1 = vbChecked Then
txtpassword.PasswordChar = ""
Else
txtpassword.PasswordChar = "*"
End If
End Sub

Private Sub cmdcancel_Click()
subclear
End Sub

Private Sub cmdlogin_Click()
    If Fnvalidatiion Then
      sublogin
     
    Else
       MsgBox "Invalid Username and Password", vbOKOnly + vbInformation, "Warning"
    End If
End Sub
Private Sub Form_Load()
    subclear
    Me.Left = 6500
    Me.Top = 3200
End Sub
