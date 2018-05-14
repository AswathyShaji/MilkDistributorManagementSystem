VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_MTYPE 
   Caption         =   "MILK TYPES"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11595
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
   ScaleHeight     =   6390
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
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
      Height          =   975
      Left            =   0
      TabIndex        =   16
      Top             =   4320
      Width           =   4335
      Begin VB.CommandButton cmddelete 
         Caption         =   "DELETE"
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
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
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
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   975
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
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
      Height          =   5295
      Left            =   4320
      TabIndex        =   14
      Top             =   0
      Width           =   6135
      Begin MSFlexGridLib.MSFlexGrid gridmtype 
         Height          =   4215
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7435
         _Version        =   393216
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW MILK TYPES"
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
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   3330
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
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtprice 
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
         TabIndex        =   3
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtcost 
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
         TabIndex        =   2
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtmtype 
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
         TabIndex        =   1
         Top             =   1800
         Width           =   1935
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
         TabIndex        =   18
         Top             =   1200
         Width           =   1740
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
         TabIndex        =   17
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label lblprice 
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
         Left            =   3600
         TabIndex        =   13
         Top             =   3000
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblcost 
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
         Left            =   3600
         TabIndex        =   12
         Top             =   2280
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblmtype 
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
         Left            =   3600
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MARKET PRICE"
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
         TabIndex        =   10
         Top             =   3360
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COST OF MILK"
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
         TabIndex        =   9
         Top             =   2640
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MILK TYPE"
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
         TabIndex        =   8
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MILK TYPES"
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
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   2235
      End
   End
End
Attribute VB_Name = "FRM_MTYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim STRSQL As String

Public Sub subinsert()
STRSQL = " INSERT INTO TBL_MILKTYPE (MT_NAME,MT_COST,MT_PRICE) " _
         & " VALUES ('" & txtmtype.Text & "' , '" & txtcost.Text & "' , '" & txtprice.Text & "' )"

Set RS = adocn.Execute(STRSQL)

End Sub
Public Sub subClear()
txtmtype.Text = ""
txtcost.Text = ""
txtprice.Text = ""
End Sub
Private Sub cmdadd_Click()
If fnValidation = True Then
subinsert
            MsgBox "Success"
           subClear
           subclearlabel
           subAddToGrid

            Else
        MsgBox "Failed", vbCritical
    End If
End Sub
Public Function fnValidation()
Dim ok1, ok2, ok3 As Boolean
If Trim(txtmtype.Text) = "" Then
 lblmtype.Visible = True
 ok1 = False
 Else
 lblmtype.Visible = False
 ok1 = True
 End If
 
 If (Not IsNumeric(txtcost.Text)) Then
 lblcost.Visible = True
 ok2 = False
 Else
 lblcost.Visible = False
 ok2 = True
 End If
 
 If (Not IsNumeric(txtprice.Text)) Then
 lblprice.Visible = True
 ok3 = False
 Else
 lblprice.Visible = False
 ok3 = True
 End If
 
If (ok1 And ok2 And ok3) = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function

Private Sub Form_Load()
subAddToGrid
subid
cmdupdate.Enabled = False
cmddelete.Enabled = False
End Sub

Private Sub txtmtype_Change()
If Trim(txtmtype.Text) = "" Then
    lblmtype.Visible = True
    Else
    lblmtype.Visible = False
End If
End Sub

Private Sub txtcost_Change()
If Trim(txtcost.Text) = "" Then
    lblcost.Visible = True
    Else
    lblcost.Visible = False
End If
End Sub
Private Sub txtprice_Change()
If Trim(txtprice.Text) = "" Then
    lblprice.Visible = True
    Else
    lblprice.Visible = False
End If
End Sub

Private Sub subSetgrid()
    gridmtype.Cols = 5
    gridmtype.Rows = 2
    gridmtype.FixedRows = 1
    gridmtype.TextMatrix(0, 1) = "SL No"
    gridmtype.TextMatrix(0, 2) = "Milk Type"
    gridmtype.TextMatrix(0, 3) = "Cost"
    gridmtype.TextMatrix(0, 4) = "Market Price"
    gridmtype.ColWidth(0) = 0
    gridmtype.ColWidth(1) = 750
    gridmtype.ColWidth(2) = 780
    gridmtype.ColWidth(3) = 780
    gridmtype.ColWidth(4) = 1600
End Sub

Public Sub subAddToGrid()
    gridmtype.Clear
    subSetgrid
    STRSQL = "select * from TBL_MILKTYPE"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridmtype.TextMatrix(i, 0) = RS!MT_ID
            gridmtype.TextMatrix(i, 1) = SLNO
            gridmtype.TextMatrix(i, 2) = RS!MT_NAME
            gridmtype.TextMatrix(i, 3) = RS!MT_COST
            gridmtype.TextMatrix(i, 4) = RS!MT_PRICE
            gridmtype.Rows = gridmtype.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridmtype.Rows = gridmtype.Rows - 1
End Sub

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_MILKTYPE"
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
lblmtype.Visible = False
lblcost.Visible = False
lblprice.Visible = False
End Sub


Private Sub gridmtype_Click()
    If gridmtype.Rows > 1 Then
        cmdupdate.Enabled = True
        cmddelete.Enabled = True
        cmdadd.Enabled = False
        STRSQL = "select * from TBL_MILKTYPE where MT_ID = '" & gridmtype.TextMatrix(gridmtype.RowSel, 0) & "'"
        Set RS = adocn.Execute(STRSQL)
        lblid.Caption = RS!MT_ID
        txtmtype.Text = RS!MT_NAME
        txtcost.Text = RS!MT_COST
        txtprice.Text = RS!MT_PRICE
    End If
End Sub

Private Sub subupdate()
 STRSQL = " UPDATE TBL_MILKTYPE SET MT_NAME= '" & txtmtype.Text & "'," _
 & " MT_COST='" & txtcost.Text & "',MT_PRICE= '" & txtprice.Text & "' where MT_ID='" & lblid.Caption & "'"
 
 Set RS = adocn.Execute(STRSQL)
End Sub

Private Sub subdelete()
STRSQL = " DELETE FROM TBL_MILKTYPE WHERE MT_ID='" & lblid.Caption & "' "
Set RS = adocn.Execute(STRSQL)
End Sub



Private Sub cmddelete_Click()
subdelete
MsgBox "deleted"
subClear
subclearlabel
subAddToGrid

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

