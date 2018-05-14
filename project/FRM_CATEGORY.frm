VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_CATEGORY 
   Caption         =   "CATEGORY DETAILS"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10485
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
   ScaleHeight     =   4980
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   11
      Top             =   4080
      Width           =   4335
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   4800
      TabIndex        =   7
      Top             =   960
      Width           =   5535
      Begin MSFlexGridLib.MSFlexGrid gridcategory 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5106
         _Version        =   393216
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW DETAILS"
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
         TabIndex        =   10
         Top             =   240
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4455
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
         Left            =   1440
         TabIndex        =   1
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lblid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   1680
         TabIndex        =   13
         Top             =   1200
         Width           =   1980
      End
      Begin VB.Label Label5 
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
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   210
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
         Left            =   3120
         TabIndex        =   6
         Top             =   1680
         Visible         =   0   'False
         Width           =   630
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
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORY OF CATTLE FEED"
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
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CATTLE FEED DETAILS"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   6120
   End
End
Attribute VB_Name = "FRM_CATEGORY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim STRSQL As String

Public Sub subinsert()
STRSQL = " INSERT INTO TBL_FEEDCATEGORY (C_NAME) VALUES ('" & txtname.Text & "')"
Set RS = adocn.Execute(STRSQL)
End Sub
Public Sub subClear()
txtname.Text = ""
End Sub
Private Sub cmdadd_Click()
If fnValidation = True Then
subinsert
subid
    MsgBox "Success"
       subClear
       subclearlabel
       subAddToGrid
            Else
            MsgBox "Registration Failed", vbCritical
            End If
           
End Sub

Public Function fnValidation()
Dim ok As Boolean
If Trim(txtname.Text) = "" Then
 lblname.Visible = True
 ok = False
 Else
 lblname.Visible = False
 ok = True
 End If
 
If ok = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function

Private Sub Form_Load()
subAddToGrid
subid
cmdupdate.Enabled = False
End Sub

Private Sub txtname_Change()
If Trim(txtname.Text) = "" Then
    lblname.Visible = True
    Else
    lblname.Visible = False
End If
End Sub

Private Sub gridcategory_Click()
    If gridcategory.Rows > 1 Then
        cmdupdate.Enabled = True
        cmdadd.Enabled = False
        STRSQL = "select * from TBL_FEEDCATEGORY where C_ID = '" & gridcategory.TextMatrix(gridcategory.RowSel, 0) & "'"
        Set RS = adocn.Execute(STRSQL)
        lblid.Caption = RS!C_ID
        txtname.Text = RS!C_NAME
    End If
End Sub

Private Sub subSetgrid()
    gridcategory.Cols = 3
    gridcategory.Rows = 2
    gridcategory.FixedRows = 1
    gridcategory.TextMatrix(0, 1) = "SL No"
    gridcategory.TextMatrix(0, 2) = "Name"
    gridcategory.ColWidth(0) = 0
    gridcategory.ColWidth(1) = 750
    gridcategory.ColWidth(2) = 1730
End Sub

Public Sub subAddToGrid()
    gridcategory.Clear
    subSetgrid
    STRSQL = "select * from TBL_FEEDCATEGORY"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridcategory.TextMatrix(i, 0) = RS!C_ID
            gridcategory.TextMatrix(i, 1) = SLNO
            gridcategory.TextMatrix(i, 2) = RS!C_NAME
            gridcategory.Rows = gridcategory.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridcategory.Rows = gridcategory.Rows - 1
End Sub

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_FEEDCATEGORY"
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
End Sub

Private Sub subupdate()
 STRSQL = " UPDATE TBL_FEEDCATEGORY SET C_NAME= '" & txtname.Text & "' where C_ID=' " & lblid.Caption & "'"
 
 Set RS = adocn.Execute(STRSQL)
End Sub


Private Sub cmdupdate_Click()
If fnValidation = True Then
subupdate
subid
MsgBox "Updation Succesfull"
subClear
subclearlabel
cmdupdate.Enabled = False
cmdadd.Enabled = True
subAddToGrid
           Else
            MsgBox "Updation Failed", vbCritical
            End If
End Sub

