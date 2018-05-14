VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_FEEDSALE 
   Caption         =   "FEED SALE"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15420
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
   ScaleHeight     =   7800
   ScaleWidth      =   15420
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
      Height          =   7215
      Left            =   6480
      TabIndex        =   20
      Top             =   360
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid gridfeedsale 
         Height          =   4815
         Left            =   360
         TabIndex        =   21
         Top             =   960
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8493
         _Version        =   393216
      End
      Begin VB.Label Label9 
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
         Left            =   720
         TabIndex        =   22
         Top             =   480
         Width           =   2670
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
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   6600
      Width           =   6375
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
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
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
         Left            =   3600
         TabIndex        =   17
         Top             =   240
         Width           =   975
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
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   6375
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
         Left            =   3120
         TabIndex        =   29
         Top             =   3960
         Width           =   1695
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
         Left            =   4920
         TabIndex        =   26
         Top             =   4680
         Width           =   975
      End
      Begin VB.ComboBox combocname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Text            =   ".........select................."
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox combocategory 
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
         Left            =   3120
         TabIndex        =   1
         Text            =   ".........select................."
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ComboBox combofid 
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
         Left            =   3120
         TabIndex        =   0
         Text            =   ".........select................."
         Top             =   1560
         Width           =   1575
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
         Left            =   3120
         TabIndex        =   13
         Top             =   5520
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
         Left            =   3120
         TabIndex        =   3
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AVAILABLE STOCK"
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
         TabIndex        =   28
         Top             =   4080
         Width           =   1890
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
         Left            =   4680
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_cost 
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
         Left            =   3120
         TabIndex        =   25
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblcid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   24
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME OF CATTLE FEED"
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
         TabIndex        =   23
         Top             =   2760
         Width           =   2355
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
         Left            =   3120
         TabIndex        =   19
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   18
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
         Left            =   4080
         TabIndex        =   15
         Top             =   5280
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
         Left            =   4200
         TabIndex        =   14
         Top             =   4440
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL COST"
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
         TabIndex        =   12
         Top             =   5520
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
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
         Top             =   4800
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRICE OF CATTLE FEED"
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
         Top             =   3360
         Width           =   2340
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORY OF CATTLE FEED"
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
         TabIndex        =   9
         Top             =   2160
         Width           =   2925
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FARMER ID"
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
         TabIndex        =   8
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FEED SALE DETAILS"
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
         Left            =   600
         TabIndex        =   6
         Top             =   120
         Width           =   3675
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   480
   End
End
Attribute VB_Name = "FRM_FEEDSALE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RS As New Recordset
Dim STRSQL As String
Dim autIn As Integer
Private Sub SUBFARMER()
STRSQL = "SELECT * FROM TBL_FARMERINF "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combofid.AddItem (RS!F_ID)
RS.MoveNext
Loop
End Sub

Private Sub SUBCATEGORY()
STRSQL = "SELECT * FROM TBL_FEEDCATEGORY "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combocategory.AddItem (RS!C_NAME)
RS.MoveNext
Loop
End Sub
Private Sub SUBFEED()
STRSQL = "SELECT * FROM TBL_FEED WHERE CF_NAME='" & combocname.List(combocname.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount = 1 Then
lbl_cost.Caption = RS!CF_PRICE
End If
End Sub


Private Sub cmdcalculate_Click()
txtcost.Text = Val(txtquantity.Text) * Val(lbl_cost.Caption)
End Sub

Private Sub cmdcancel_Click()
subClear
End Sub


Private Sub combocategory_Click()
'STRSQL = "select C_ID from TBL_FEEDCATEGORY where C_NAME='" & combocategory.List(combocategory.ListIndex) & "'"
'Set RS = adocn.Execute(STRSQL)
'If RS.RecordCount = 1 Then
'lblcid.Caption = RS!C_ID
'End If
fillname
End Sub

Private Sub combocname_Click()
SUBFEED
substock
End Sub

Private Sub Form_Load()
'SUBFEEDNAME
SUBFARMER
SUBCATEGORY
SUBFEED
subAddToGrid
subid
lbldate.Caption = DateValue(Now)
End Sub

Public Sub subinsert()

STRSQL = " INSERT INTO TBL_FEEDSALE (F_ID,C_NAME,CF_NAME,CF_PRICE,S_COST,S_QUANTITY,S_DATE) " _
          & " VALUES ('" & combofid.List(combofid.ListIndex) & "','" & combocategory.List(combocategory.ListIndex) & "'," _
          & " '" & combocname.List(combocname.ListIndex) & "','" & lbl_cost.Caption & "'," _
          & " '" & txtcost.Text & " ','" & txtquantity.Text & " ','" & lbldate.Caption & "')"
Set RS = adocn.Execute(STRSQL)
End Sub

Public Sub subClear()
txtcost.Text = ""
txtquantity.Text = ""
End Sub
Private Sub cmdadd_Click()
If combofid.Text = ".........select................." Or combofid.Text = "" Then
 MsgBox "select the farmer id"
 Else
 If combocategory.Text = ".........select................." Or combocategory.Text = "" Then
 MsgBox "select the category"
 Else
 If combocname.Text = ".........select................." Or combocname.Text = "" Then
 MsgBox "select the cattle feed"
 Else
If fnValidation = True Then
subinsert
substockupdation
MsgBox "Success"
subClear
subclearlabel
subAddToGrid
    Else
        MsgBox "Registration Failed", vbCritical
    End If
    End If
    End If
    End If
End Sub

Public Function fnValidation()
Dim ok1, ok2 As Boolean
If Trim(txtquantity.Text) = "" Then
 lblquantity.Visible = True
 ok1 = False
 Else
 lblquantity.Visible = False
 ok1 = True
 End If
 
 If (Not IsNumeric(txtcost.Text)) Then
 lblcost.Visible = True
 ok2 = False
 Else
 lblcost.Visible = False
 ok2 = True
 End If
 
If (ok1 And ok2) = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function

'Private Sub txtquantity_Change()
'If Trim(txtquantity.Text) = "" Then
'    lblquantity.Visible = True
'
'    Else
'    lblquantity.Visible = False
'End If
' txtcost.Text = ""
'End Sub

Private Sub txtcost_Change()
If Trim(txtcost.Text) = "" Then
    lblcost.Visible = True
    Else
    lblcost.Visible = False
End If
End Sub

Private Sub subSetgrid()
    gridfeedsale.Cols = 6
    gridfeedsale.Rows = 2
    gridfeedsale.FixedRows = 1
    gridfeedsale.TextMatrix(0, 1) = "SL No"
    gridfeedsale.TextMatrix(0, 2) = "Cattle feed name"
    gridfeedsale.TextMatrix(0, 3) = "Price"
    gridfeedsale.TextMatrix(0, 4) = "Quantity"
    gridfeedsale.TextMatrix(0, 5) = "Total cost"
    gridfeedsale.ColWidth(0) = 0
    gridfeedsale.ColWidth(1) = 750
    gridfeedsale.ColWidth(2) = 1730
    gridfeedsale.ColWidth(3) = 1730
    gridfeedsale.ColWidth(4) = 1600
    gridfeedsale.ColWidth(5) = 2000
End Sub

Public Sub subAddToGrid()
    gridfeedsale.Clear
    subSetgrid
    STRSQL = "select * from TBL_FEEDSALE"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridfeedsale.TextMatrix(i, 0) = RS!S_ID
            gridfeedsale.TextMatrix(i, 1) = SLNO
            gridfeedsale.TextMatrix(i, 2) = RS!CF_NAME
            gridfeedsale.TextMatrix(i, 3) = RS!CF_PRICE
            gridfeedsale.TextMatrix(i, 4) = RS!S_QUANTITY
            gridfeedsale.TextMatrix(i, 5) = RS!S_COST
            gridfeedsale.Rows = gridfeedsale.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridfeedsale.Rows = gridfeedsale.Rows - 1
End Sub

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
   STRSQL = "select * from TBL_FEEDSALE"
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
lblcost.Visible = False
End Sub

Public Function fillname()
STRSQL = "SELECT * FROM TBL_FEED where C_NAME='" & combocategory.List(combocategory.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combocname.AddItem (RS!CF_NAME)
RS.MoveNext
Loop
End Function


Private Sub substock()
STRSQL = "SELECT * FROM TBL_FEED where CF_NAME='" & combocname.List(combocname.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount = 0 Then
txtstock.Text = "0"
Else
txtstock = RS!CF_QUANTITY
End If
End Sub
Private Sub substockupdation()
TOTAL_QUANTITY = Val(txtstock.Text) - Val(txtquantity.Text)
STRSQL = " UPDATE TBL_FEED SET CF_QUANTITY= '" & TOTAL_QUANTITY & "' " _
    & " where CF_NAME='" & combocname.List(combocname.ListIndex) & "'"
 Set RS = adocn.Execute(STRSQL)
End Sub


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
Private Sub combocategory_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub
Private Sub combocname_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub
Private Sub combofid_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub
