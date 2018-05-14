VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRM_STOCK 
   Caption         =   "STOCK UPDATION"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      Begin VB.TextBox txtquantity 
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
         Left            =   2400
         TabIndex        =   2
         Top             =   2880
         Width           =   1695
      End
      Begin VB.ComboBox combomtype 
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
         Left            =   2400
         TabIndex        =   1
         Top             =   2160
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   33751041
         CurrentDate     =   42600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COLLECTION DATE"
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
         Top             =   1440
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK DETAILS"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   2400
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MILK TYPE "
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY OF MILK"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblquantity 
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
         Left            =   3480
         TabIndex        =   5
         Top             =   2640
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   210
      End
      Begin VB.Label lblid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Width           =   1380
      End
   End
End
Attribute VB_Name = "FRM_STOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New Recordset
Dim STRSQL As String
Dim autIn As Integer
Private Sub SUBMTYPE()
STRSQL = "SELECT * FROM TBL_MILKTYPE "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combomtype.AddItem (RS!MT_NAME)
RS.MoveNext
Loop
End Sub

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_STOCK"
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

Public Function fnValidation()
Dim ok1 As Boolean

 If (Not IsNumeric(txtquantity.Text)) Then
 lblquantity.Visible = True
 ok1 = False
 Else
 lblquantity.Visible = False
 ok1 = True
 End If
 
If ok1 = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function

Private Sub combomtype_Click()
SUBQUANTITY
End Sub

Private Sub Form_Load()
SUBMTYPE
'lbldate.Caption = DateValue(Now)

End Sub

Private Sub txtquantity_Change()
If Trim(txtquantity.Text) = "" Then
    lblquantity.Visible = True
    Else
    lblquantity.Visible = False
End If
End Sub

Private Sub subclearlabel()
lblquantity.Visible = False
End Sub

Private Sub SUBQUANTITY()
STRSQL = "SELECT * FROM TBL_MCOLLECTION where MT_NAME='" & combomtype.List(combomtype.ListIndex) & "' AND COLLECTIONDATE='" & DTPicker1 & "' "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
txtquantity.Text = RS!M_QUANTITY
RS.MoveNext
Loop
End Sub

