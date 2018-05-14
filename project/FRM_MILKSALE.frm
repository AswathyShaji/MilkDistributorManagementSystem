VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_MILKSALE 
   Caption         =   "MILK SALE"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14400
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
   ScaleHeight     =   5745
   ScaleWidth      =   14400
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
      Height          =   5775
      Left            =   4920
      TabIndex        =   16
      Top             =   0
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid gridmsale 
         Height          =   4335
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7646
         _Version        =   393216
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW DETAILS"
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
         Left            =   480
         TabIndex        =   18
         Top             =   480
         Width           =   2670
      End
   End
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
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
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
         Left            =   3600
         TabIndex        =   22
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox txtstock 
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
         Left            =   2280
         TabIndex        =   21
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdcalculate 
         Caption         =   "calculate"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtcost 
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
         Left            =   2280
         TabIndex        =   13
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox txtprice 
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
         Left            =   2280
         TabIndex        =   2
         Top             =   3000
         Width           =   1455
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
         Left            =   2280
         TabIndex        =   1
         Text            =   ".........select................."
         Top             =   1440
         Width           =   1575
      End
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
         Left            =   2280
         TabIndex        =   3
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AVAILABLE STOCK"
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
         TabIndex        =   20
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Label lbldate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   480
         Width           =   1335
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
         TabIndex        =   15
         Top             =   960
         Width           =   1380
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
         TabIndex        =   14
         Top             =   960
         Width           =   210
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
         Left            =   3120
         TabIndex        =   12
         Top             =   4320
         Visible         =   0   'False
         Width           =   630
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
         Left            =   3120
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   630
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
         Left            =   3120
         TabIndex        =   10
         Top             =   3600
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL COST"
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
         TabIndex        =   9
         Top             =   4680
         Width           =   1230
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   1410
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
         TabIndex        =   7
         Top             =   3960
         Width           =   1815
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MILK SALE DETAILS"
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3150
      End
   End
End
Attribute VB_Name = "FRM_MILKSALE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RS As New Recordset
Dim RS1 As New Recordset
Dim STRSQL As String
Dim autIn As Integer
Dim STRSQL1 As String

Private Sub SUBMTYPE()
STRSQL = "SELECT * FROM TBL_MILKTYPE "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combomtype.AddItem (RS!MT_NAME)
RS.MoveNext
Loop
End Sub

Private Sub cmdcalculate_Click()
txtcost.Text = Val(txtquantity.Text) * Val(txtprice.Text)
End Sub

Private Sub combomtype_Click()
SUBPRICE
substock
End Sub

Private Sub Form_Load()
SUBMTYPE
subAddToGrid
subid
lbldate.Caption = DateValue(Now)
txtstock.Enabled = False
End Sub

Public Sub subinsert()

STRSQL = " INSERT INTO TBL_MILKSALE (MT_NAME,MS_DATE,MS_QUANTITY,MT_PRICE,MS_COST) " _
          & " VALUES ('" & combomtype.List(combomtype.ListIndex) & "','" & lbldate.Caption & "'," _
          & " '" & txtquantity.Text & "' ,'" & txtprice.Text & "', '" & txtcost.Text & " ')"
Set RS = adocn.Execute(STRSQL)
End Sub

Public Sub subClear()
txtquantity.Text = ""
txtcost.Text = ""
txtprice.Text = ""
End Sub
Private Sub cmdadd_Click()
If combomtype.Text = ".........select................." Or combomtype.Text = "" Then
 MsgBox "select the milktype"
 Else
If fnValidation = True Then
subinsert
substockupdation
MsgBox "Success"
subClear
subclearlabel
subAddToGrid
    Else
        MsgBox "Failed", vbCritical
    End If
    End If
End Sub
Public Function fnValidation()
Dim ok2, ok3, ok4 As Boolean

 
 If (Not IsNumeric(txtquantity.Text)) Then
 lblquantity.Visible = True
 ok2 = False
 Else
 lblquantity.Visible = False
 ok2 = True
 End If
 
 If (Not IsNumeric(txtprice.Text)) Then
 lblprice.Visible = True
 ok3 = False
 Else
 lblprice.Visible = False
 ok3 = True
 End If
 
 If (Not IsNumeric(txtcost.Text)) Then
 lblcost.Visible = True
 ok4 = False
 Else
 lblcost.Visible = False
 ok4 = True
 End If
 
If (ok2 And ok3 And ok4) = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function

Private Sub txtquantity_Change()
If Val(txtquantity.Text) > Val(txtstock.Text) Then
cmdcalculate.Enabled = False
cmdadd.Enabled = False
MsgBox "required quantity not available"
Else
cmdcalculate.Enabled = True
cmdadd.Enabled = True
End If
If Trim(txtquantity.Text) = "" Then
    lblquantity.Visible = True
    Else
    lblquantity.Visible = False
End If
End Sub

Private Sub txtprice_Change()
If Trim(txtprice.Text) = "" Then
    lblprice.Visible = True
    Else
    lblprice.Visible = False
End If
End Sub
Private Sub txtcost_Change()
If Trim(txtcost.Text) = "" Then
    lblcost.Visible = True
    Else
    lblcost.Visible = False
End If
End Sub

Private Sub subSetgrid()
    gridmsale.Cols = 6
    gridmsale.Rows = 2
    gridmsale.FixedRows = 1
    gridmsale.TextMatrix(0, 1) = "SL No"
    gridmsale.TextMatrix(0, 2) = "Milk Type"
    gridmsale.TextMatrix(0, 3) = "Date Of Sale"
    gridmsale.TextMatrix(0, 4) = "Quantity"
    gridmsale.TextMatrix(0, 5) = "Total cost"
    gridmsale.ColWidth(0) = 0
    gridmsale.ColWidth(1) = 750
    gridmsale.ColWidth(2) = 1730
    gridmsale.ColWidth(3) = 1730
    gridmsale.ColWidth(4) = 1600
    gridmsale.ColWidth(5) = 2000
End Sub

Public Sub subAddToGrid()
    gridmsale.Clear
    subSetgrid
    STRSQL = "select * from TBL_MILKSALE"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridmsale.TextMatrix(i, 0) = RS!MS_ID
            gridmsale.TextMatrix(i, 1) = SLNO
            gridmsale.TextMatrix(i, 2) = RS!MT_NAME
            gridmsale.TextMatrix(i, 3) = RS!MS_DATE
            gridmsale.TextMatrix(i, 4) = RS!MS_QUANTITY
            gridmsale.TextMatrix(i, 5) = RS!MS_COST
            gridmsale.Rows = gridmsale.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridmsale.Rows = gridmsale.Rows - 1
End Sub

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_MILKSALE"
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
lblquantity.Visible = False
lblprice.Visible = False
lblcost.Visible = False
End Sub

Public Function fillprice()
STRSQL = "SELECT * FROM TBL_MILKTYPE where MT_NAME='" & combomtype.List(combomtype.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
txtprice.Text = RS!MT_PRICE
RS.MoveNext
Loop
End Function

Private Sub SUBPRICE()
STRSQL = "SELECT * FROM TBL_MILKTYPE WHERE MT_NAME='" & combomtype.List(combomtype.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount = 1 Then
txtprice.Text = RS!MT_PRICE
End If
End Sub

Private Sub substock()
STRSQL = "SELECT * FROM TBL_STOCK where MT_NAME='" & combomtype.List(combomtype.ListIndex) & "'" _
           & " AND COLLECTIONDATE='" & lbldate.Caption & "' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount = 0 Then
txtstock.Text = "0"
Else
txtstock = RS!M_QUANTITY
End If
End Sub
Private Sub substockupdation()
TOTAL_QUANTITY = Val(txtstock.Text) - Val(txtquantity.Text)
STRSQL = " UPDATE TBL_STOCK SET M_QUANTITY= '" & TOTAL_QUANTITY & "' " _
    & " where MT_NAME='" & combomtype.List(combomtype.ListIndex) & "'" _
    & " AND COLLECTIONDATE='" & lbldate.Caption & "' "
 Set RS = adocn.Execute(STRSQL)
End Sub
Private Sub combomtype_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub
