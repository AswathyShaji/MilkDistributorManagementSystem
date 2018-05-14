VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRM_FARMERPAYMENT 
   Caption         =   "PAYMENT DETAILS"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14865
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
   ScaleHeight     =   8085
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000E&
      Caption         =   "Feed sale details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   6000
      TabIndex        =   24
      Top             =   3720
      Width           =   8655
      Begin MSFlexGridLib.MSFlexGrid gridfeed 
         Height          =   2055
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   3625
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Milk collection details"
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
      Left            =   6000
      TabIndex        =   23
      Top             =   480
      Width           =   8535
      Begin MSFlexGridLib.MSFlexGrid gridfarmer 
         Height          =   2055
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3625
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   360
      TabIndex        =   21
      Top             =   6720
      Width           =   5535
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
         Left            =   3120
         TabIndex        =   22
         Top             =   120
         Width           =   1095
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
      Height          =   6495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      Begin VB.CommandButton cmdcalculate 
         Caption         =   "calculate"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   840
         TabIndex        =   27
         Top             =   2880
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109445121
         CurrentDate     =   42629
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109445121
         CurrentDate     =   42629
      End
      Begin VB.ComboBox combofid 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         TabIndex        =   3
         Text            =   ".....select........"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtbalance 
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
         Left            =   2640
         TabIndex        =   16
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox txttp 
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
         Left            =   2640
         TabIndex        =   15
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtfp 
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
         Left            =   2640
         TabIndex        =   14
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtmp 
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
         Left            =   2640
         TabIndex        =   13
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblbalance 
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
         TabIndex        =   20
         Top             =   5640
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lbltp 
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
         TabIndex        =   19
         Top             =   4920
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblfp 
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
         TabIndex        =   18
         Top             =   4200
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblmp 
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
         TabIndex        =   17
         Top             =   3480
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BALANCE"
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
         TabIndex        =   12
         Top             =   5880
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FARMER ID"
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
         TabIndex        =   11
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAYMENT"
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
         TabIndex        =   10
         Top             =   5160
         Width           =   1590
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL FEED PRICE"
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
         Top             =   4560
         Width           =   1785
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL MILK PAYMENT"
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
         Top             =   3840
         Width           =   2100
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TO DATE"
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
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FROM DATE"
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
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAYMENT DETAILS"
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
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
   End
End
Attribute VB_Name = "FRM_FARMERPAYMENT"
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

Private Sub cmdcalculate_Click()
submilkpayment
subfeedpayment
subcalculate
End Sub

Private Sub combofid_Click()
subAddToGridfeed
subAddToGrid
End Sub


Private Sub DTPicker1_Click()
subAddToGrid
subAddToGridfeed
End Sub

Private Sub DTPicker2_Click()
subAddToGrid
subAddToGridfeed
End Sub

Private Sub Form_Load()
SUBFARMER
End Sub

Public Sub subinsert()
STRSQL = " INSERT INTO TBL_PAYMENT (F_ID,FROMDATE,TODATE,P_MILK,P_FEED,P_PAYMENT,P_BALANCE) " _
          & " VALUES ('" & combofid.List(combofid.ListIndex) & "','" & DTPicker1 & "' ," _
          & " '" & DTPicker2 & "' ,'" & txtmp.Text & "','" & txtfp.Text & "','" & txttp.Text & "'," _
          & " '" & txtbalance.Text & "')"
Set RS = adocn.Execute(STRSQL)
End Sub

Public Sub subClear()
txtbalance.Text = ""
txtfp.Text = ""
txtmp.Text = ""
txttp.Text = ""
End Sub

Private Sub cmdadd_Click()
If combofid.Text = ".....select........" Or combofid.Text = "" Then
 MsgBox "select the milktype"
 Else
If fnValidation = True Then
subinsert
MsgBox "Success"
subClear
subclearlabel
    Else
        MsgBox "Registration Failed", vbCritical
    End If
    End If
End Sub

Public Function fnValidation()
Dim ok1, ok2, ok3, ok4, ok5, ok6 As Boolean
  
 If (Not IsNumeric(txtmp.Text)) Then
 lblmp.Visible = True
 ok1 = False
 Else
 lblmp.Visible = False
 ok1 = True
 End If
 
 If (Not IsNumeric(txtfp.Text)) Then
 lblfp.Visible = True
 ok2 = False
 Else
 lblfp.Visible = False
 ok2 = True
 End If
 
 If (Not IsNumeric(txttp.Text)) Then
 lbltp.Visible = True
 ok3 = False
 Else
 lbltp.Visible = False
 ok3 = True
 End If
 
 If (Not IsNumeric(txtbalance.Text)) Then
 lblbalance.Visible = True
 ok4 = False
 Else
 lblbalance.Visible = False
 ok4 = True
 End If
 
If (ok1 And ok2 And ok3 And ok4) = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function

Private Sub txtmp_Change()
If Trim(txtmp.Text) = "" Then
    lblmp.Visible = True
    Else
    lblmp.Visible = False
End If
End Sub

Private Sub txtfp_Change()
If Trim(txtfp.Text) = "" Then
    lblfp.Visible = True
    Else
    lblfp.Visible = False
End If
End Sub

Private Sub txttp_Change()
If Trim(txttp.Text) = "" Then
    lbltp.Visible = True
    Else
    lbltp.Visible = False
End If
End Sub

Private Sub txtbalance_Change()
If Trim(txtbalance.Text) = "" Then
    lblbalance.Visible = True
    Else
    lblbalance.Visible = False
End If
End Sub

Private Sub subclearlabel()
lblbalance.Visible = False
lblfp.Visible = False
lblmp.Visible = False
lbltp.Visible = False
End Sub

Private Sub subSetgrid()
    gridfarmer.Cols = 5
    gridfarmer.Rows = 2
    gridfarmer.FixedRows = 1
    gridfarmer.TextMatrix(0, 1) = "SL No"
    gridfarmer.TextMatrix(0, 2) = "Quantity"
    gridfarmer.TextMatrix(0, 3) = "Collection date"
    gridfarmer.TextMatrix(0, 4) = "Cost"
    gridfarmer.ColWidth(0) = 0
    gridfarmer.ColWidth(1) = 750
    gridfarmer.ColWidth(2) = 1730
    gridfarmer.ColWidth(3) = 1730
    gridfarmer.ColWidth(4) = 1600
End Sub

Public Sub subAddToGrid()
    gridfarmer.Clear
    subSetgrid
    STRSQL = "select * from TBL_MCOLLECTION where F_ID='" & combofid.List(combofid.ListIndex) & "' and COLLECTIONDATE between '" & DTPicker1 & "' and '" & DTPicker2 & "'"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridfarmer.TextMatrix(i, 0) = RS!M_ID
            gridfarmer.TextMatrix(i, 1) = SLNO
            gridfarmer.TextMatrix(i, 2) = RS!M_QUANTITY
            gridfarmer.TextMatrix(i, 3) = RS!COLLECTIONDATE
            gridfarmer.TextMatrix(i, 4) = RS!TOTALCOST
            gridfarmer.Rows = gridfarmer.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridfarmer.Rows = gridfarmer.Rows - 1
End Sub

Private Sub subSetgridfeed()
    gridfeed.Cols = 5
    gridfeed.Rows = 2
    gridfeed.FixedRows = 1
    gridfeed.TextMatrix(0, 1) = "SL No"
    gridfeed.TextMatrix(0, 2) = "DATE"
    gridfeed.TextMatrix(0, 3) = "QUANTITY"
    gridfeed.TextMatrix(0, 4) = "PRICE"
    gridfeed.ColWidth(0) = 0
    gridfeed.ColWidth(1) = 750
    gridfeed.ColWidth(2) = 1730
    gridfeed.ColWidth(3) = 1730
    gridfeed.ColWidth(4) = 1600
End Sub

Public Sub subAddToGridfeed()
    gridfeed.Clear
    subSetgridfeed
    STRSQL = "select * from TBL_FEEDSALE where F_ID='" & combofid.List(combofid.ListIndex) & "' and S_DATE between '" & DTPicker1 & "' and '" & DTPicker2 & "'"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridfeed.TextMatrix(i, 0) = RS!S_ID
            gridfeed.TextMatrix(i, 1) = SLNO
            gridfeed.TextMatrix(i, 2) = RS!S_DATE
            gridfeed.TextMatrix(i, 3) = RS!S_QUANTITY
            gridfeed.TextMatrix(i, 4) = RS!S_COST
            gridfeed.Rows = gridfeed.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridfeed.Rows = gridfeed.Rows - 1
End Sub


Private Sub submilkpayment()
Dim price As String
STRSQL = "select sum(TOTALCOST) AS PRICE from TBL_MCOLLECTION " _
   & " WHERE F_ID='" & combofid.List(combofid.ListIndex) & "' and" _
   & " COLLECTIONDATE between '" & DTPicker1 & "' and '" & DTPicker2 & "'"
Set RS = adocn.Execute(STRSQL)
txtmp.Text = RS!price
End Sub

Private Sub subfeedpayment()
Dim price As String
STRSQL = "select sum(S_COST) as price from TBL_FEEDSALE " _
   & " WHERE F_ID='" & combofid.List(combofid.ListIndex) & "' and" _
   & " S_DATE between '" & DTPicker1 & "' and '" & DTPicker2 & "'"
Set RS = adocn.Execute(STRSQL)
txtfp.Text = RS!price
End Sub

Private Sub subcalculate()
Dim mp As Double
Dim fp As Double
Dim tp As Double
mp = Val(txtmp.Text)
fp = Val(txtfp.Text)
If mp > fp Then
tp = mp - fp
txttp.Text = tp
txtbalance.Text = "0"
Else
tp = fp - mp
txtbalance.Text = tp
txttp.Text = "0"
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
