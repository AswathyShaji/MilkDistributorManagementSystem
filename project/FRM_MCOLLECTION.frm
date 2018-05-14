VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRM_MCOLLECTION 
   Caption         =   "DAILY MILK COLLECTION"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
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
   ScaleHeight     =   8385
   ScaleWidth      =   14535
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
      Height          =   7335
      Left            =   5880
      TabIndex        =   14
      Top             =   0
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid gridmcollection 
         Height          =   4935
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8705
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
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   2670
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
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdadd 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4200
         TabIndex        =   26
         Top             =   6720
         Width           =   1095
      End
      Begin VB.CommandButton cmdcalculate 
         Caption         =   "CALCULATE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   25
         Top             =   5160
         Width           =   1215
      End
      Begin VB.ComboBox combomquality 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Text            =   ".........select................."
         Top             =   3600
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   5520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   48824321
         CurrentDate     =   42600
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
         Left            =   2880
         TabIndex        =   4
         Top             =   4800
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
         Left            =   2760
         TabIndex        =   2
         Text            =   ".........select................."
         Top             =   2400
         Width           =   1455
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
         Left            =   2880
         TabIndex        =   11
         Top             =   6240
         Width           =   1455
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
         Left            =   2760
         TabIndex        =   1
         Text            =   ".........select................."
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblcostinp 
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
         Left            =   2760
         TabIndex        =   24
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COST IN PERCENTAGE"
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
         Top             =   4200
         Width           =   2205
      End
      Begin VB.Label lblmcost 
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
         Left            =   2760
         TabIndex        =   22
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUALITY OF MILK"
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
         TabIndex        =   21
         Top             =   3600
         Width           =   1800
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COST OF MILK"
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
         TabIndex        =   20
         Top             =   3120
         Width           =   1470
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
         Left            =   2880
         TabIndex        =   16
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label7 
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
         TabIndex        =   15
         Top             =   1080
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
         Left            =   3720
         TabIndex        =   13
         Top             =   6000
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
         Left            =   3720
         TabIndex        =   12
         Top             =   4560
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label3 
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
         Top             =   5520
         Width           =   1905
      End
      Begin VB.Label Label6 
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
         TabIndex        =   9
         Top             =   6240
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY OF MILK"
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
         Top             =   4800
         Width           =   1965
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE OF MILK"
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
         TabIndex        =   7
         Top             =   2520
         Width           =   1410
      End
      Begin VB.Label Label2 
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
         TabIndex        =   6
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MILK COLLECTION DETAILS"
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5025
      End
   End
End
Attribute VB_Name = "FRM_MCOLLECTION"
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

Private Sub SUBMQUALITY()
STRSQL = "SELECT * FROM TBL_MILKTYPE "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combomtype.AddItem (RS!MT_NAME)
RS.MoveNext
Loop
End Sub
Private Sub SUBQUALITY()
STRSQL = "SELECT * FROM TBL_CHART "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combomquality.AddItem (RS!QUALITY)
RS.MoveNext
Loop
End Sub

Private Sub cmdadd_Click()
If combofid.Text = ".........select................." Or combofid.Text = "" Then
 MsgBox "select the farmer id"
 Else
 If combomtype.Text = ".........select................." Or combomtype.Text = "" Then
 MsgBox "select the milk type"
 Else
 If combomquality.Text = ".........select................." Or combomquality.Text = "" Then
 MsgBox "select the quantity of milk"
 Else
If fnValidation = True Then
subinsert
substock
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


Private Sub cmdcalculate_Click()
Dim cost As Double
cost = Val(lblmcost.Caption) * Val(lblcostinp.Caption) / 100
txtcost.Text = cost * Val(txtquantity.Text)
End Sub

Private Sub combomquality_Click()
SUBCOST
End Sub

Private Sub combomtype_Click()
SUBPRICE
End Sub

Private Sub Form_Load()
SUBFARMER
SUBMQUALITY
subAddToGrid
subid
SUBQUALITY
End Sub
Public Sub subinsert()

STRSQL = " INSERT INTO TBL_MCOLLECTION (F_ID,MT_NAME,M_QUANTITY,COLLECTIONDATE,TOTALCOST,QUALITY) " _
          & " VALUES ( '" & combofid.List(combofid.ListIndex) & "','" & combomtype.List(combomtype.ListIndex) & "' , " _
          & " '" & txtquantity.Text & "','" & DTPicker1 & "','" & txtcost.Text & "','" & combomquality.List(combomquality.ListIndex) & "')"
Set RS = adocn.Execute(STRSQL)

End Sub

Public Sub subClear()
txtquantity.Text = ""
txtcost.Text = ""
End Sub
Public Function fnValidation()
Dim ok1, ok2 As Boolean

 If (Not IsNumeric(txtquantity.Text)) Then
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

Private Sub txtquantity_Change()
If Trim(txtquantity.Text) = "" Then
    lblquantity.Visible = True
    Else
    lblquantity.Visible = False
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
    gridmcollection.Cols = 5
    gridmcollection.Rows = 2
    gridmcollection.FixedRows = 1
    gridmcollection.TextMatrix(0, 1) = "SL No"
    gridmcollection.TextMatrix(0, 2) = "Milk type"
    gridmcollection.TextMatrix(0, 3) = "quantity of milk"
    gridmcollection.TextMatrix(0, 4) = "Cost"
    gridmcollection.ColWidth(0) = 0
    gridmcollection.ColWidth(1) = 750
    gridmcollection.ColWidth(2) = 1730
    gridmcollection.ColWidth(3) = 1730
    gridmcollection.ColWidth(4) = 1600
End Sub

Public Sub subAddToGrid()
    gridmcollection.Clear
    subSetgrid
    STRSQL = "select * from TBL_MCOLLECTION"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridmcollection.TextMatrix(i, 0) = RS!M_ID
            gridmcollection.TextMatrix(i, 1) = SLNO
            gridmcollection.TextMatrix(i, 2) = RS!MT_NAME
            gridmcollection.TextMatrix(i, 3) = RS!M_QUANTITY
            gridmcollection.TextMatrix(i, 4) = RS!TOTALCOST
            gridmcollection.Rows = gridmcollection.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridmcollection.Rows = gridmcollection.Rows - 1
End Sub

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_MCOLLECTION"
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

Private Sub SUBPRICE()
STRSQL = "SELECT * FROM TBL_MILKTYPE WHERE MT_NAME='" & combomtype.List(combomtype.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount = 1 Then
lblmcost.Caption = RS!MT_PRICE
End If
End Sub

Private Sub SUBCOST()
STRSQL = "SELECT * FROM TBL_CHART WHERE QUALITY='" & combomquality.List(combomquality.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount = 1 Then
lblcostinp.Caption = RS!QL_COST
End If
End Sub

Private Sub substock()
STRSQL = "SELECT * FROM TBL_STOCK where MT_NAME='" & combomtype.List(combomtype.ListIndex) & "'" _
           & " AND COLLECTIONDATE='" & DTPicker1 & "' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount = 0 Then
STRSQL = " INSERT INTO TBL_STOCK (MT_NAME,M_QUANTITY,COLLECTIONDATE) " _
          & " VALUES ( '" & combomtype.List(combomtype.ListIndex) & "' , " _
          & " '" & txtquantity.Text & "','" & DTPicker1 & "')"
Set RS = adocn.Execute(STRSQL)
Else
Dim QNTITY As String
QNTITY = RS!M_QUANTITY
Dim TOTAL_QUANTITY As String
TOTAL_QUANTITY = Val(txtquantity.Text) + QNTITY
STRSQL = " UPDATE TBL_STOCK SET M_QUANTITY= '" & TOTAL_QUANTITY & "' " _
    & " where MT_NAME='" & combomtype.List(combomtype.ListIndex) & "'" _
    & " AND COLLECTIONDATE='" & DTPicker1 & "' "
 Set RS = adocn.Execute(STRSQL)
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
Private Sub combomquality_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub
Private Sub combomtype_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub
