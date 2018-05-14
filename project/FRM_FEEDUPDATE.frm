VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_FEEDUPDATE 
   Caption         =   "UPDATE FEED"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14580
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
   ScaleHeight     =   7620
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   8760
      TabIndex        =   3
      Top             =   240
      Width           =   4935
      Begin VB.CommandButton cmdsubmit 
         Caption         =   "SUBMIT"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton cmdustock 
         Caption         =   "UPDATE STOCK"
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
         Left            =   2520
         TabIndex        =   9
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox txtupdatequantity 
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
         Left            =   360
         TabIndex        =   8
         Top             =   3960
         Width           =   1455
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
         Left            =   360
         TabIndex        =   7
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton cmduprice 
         Caption         =   "UPDATE PRICE"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
      End
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
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtupdateprice 
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
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblsupdate 
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
         Left            =   1200
         TabIndex        =   15
         Top             =   3720
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblpupdate 
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
         Left            =   1200
         TabIndex        =   14
         Top             =   1680
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
         Left            =   1200
         TabIndex        =   13
         Top             =   960
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
         Left            =   1200
         TabIndex        =   12
         Top             =   2760
         Visible         =   0   'False
         Width           =   630
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
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1980
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
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid gridfeed 
         Height          =   3855
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6800
         _Version        =   393216
      End
      Begin VB.Label Label7 
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
         TabIndex        =   2
         Top             =   480
         Width           =   2670
      End
   End
End
Attribute VB_Name = "FRM_FEEDUPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim STRSQL As String

Public Sub subinsert()
Private Sub subSetgrid()
    gridfeed.Cols = 6
    gridfeed.Rows = 2
    gridfeed.FixedRows = 1
    gridfeed.TextMatrix(0, 1) = "SL No"
    gridfeed.TextMatrix(0, 2) = "Category"
    gridfeed.TextMatrix(0, 3) = "Name of cattle feed"
    gridfeed.TextMatrix(0, 4) = "Quantity"
    gridfeed.TextMatrix(0, 5) = "price"
    gridfeed.ColWidth(0) = 0
    gridfeed.ColWidth(1) = 750
    gridfeed.ColWidth(2) = 750
    gridfeed.ColWidth(3) = 730
    gridfeed.ColWidth(4) = 750
    gridfeed.ColWidth(5) = 750
End Sub

Public Sub subAddToGrid()
    gridfeed.Clear
    subSetgrid
    STRSQL = "select * from TBL_FEED"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridfeed.TextMatrix(i, 0) = RS!CF_ID
            gridfeed.TextMatrix(i, 1) = SLNO
            gridfeed.TextMatrix(i, 2) = RS!C_NAME
            gridfeed.TextMatrix(i, 3) = RS!CF_NAME
            gridfeed.TextMatrix(i, 4) = RS!CF_QUANTITY
            gridfeed.TextMatrix(i, 5) = RS!CF_PRICE
            gridfeed.Rows = gridfeed.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridfeed.Rows = gridfeed.Rows - 1
End Sub

Private Sub cmdsubmit_Click()
If fnValidation Then
If txtupdateprice.Text = "" Then
subupdatestock
Else
subupdateprice
subupdatestock
End If
MsgBox "Success"
'subclearlabel
Else
MsgBox "failed"

subAddToGrid
End If

txtprice.Text = ""
txtupdateprice.Text = ""
txtquantity.Text = ""
txtupdatequantity.Text = "0"
txtprice.Enabled = False
txtupdateprice.Enabled = False
txtquantity.Enabled = False
txtupdatequantity.Enabled = False
End Sub

Private Sub cmduprice_Click()
txtupdateprice.Enabled = True

End Sub

Private Sub cmdustock_Click()
subupdatestock
txtupdatequantity.Enabled = True
End Sub

Private Sub Form_Load()
txtupdatequantity.Text = 0
subAddToGrid
txtprice.Enabled = False
txtupdateprice.Enabled = False
txtquantity.Enabled = False
txtupdatequantity.Enabled = False
End Sub


Private Sub subupdateprice()
 
 STRSQL = " UPDATE TBL_FEED SET CF_PRICE= '" & txtupdateprice.Text & "'where CF_ID=' " & lblid.Caption & "'"
 Set RS = adocn.Execute(STRSQL)
End Sub

Private Sub gridfeed_Click()
    If gridfeed.Rows > 1 Then
        STRSQL = "select * from TBL_FEED where CF_ID = '" & gridfeed.TextMatrix(gridfeed.RowSel, 0) & "'"
        Set RS = adocn.Execute(STRSQL)
        lblid.Caption = RS!CF_ID
        txtprice.Text = RS!CF_PRICE
        txtquantity.Text = RS!CF_QUANTITY
    End If
End Sub
Private Sub subupdatestock()

 txtquantity.Text = Val(txtupdatequantity.Text) + Val(txtquantity.Text)
 STRSQL = " UPDATE TBL_FEED SET CF_QUANTITY= '" & txtquantity.Text & "'where CF_ID=' " & lblid.Caption & "'"
 Set RS = adocn.Execute(STRSQL)
End Sub



Public Function fnValidation()
Dim ok1, ok2, ok3, ok4 As Boolean
 
If (Not IsNumeric(txtquantity.Text)) Then
 lblquantity.Visible = True
 ok1 = False
 Else
 lblquantity.Visible = False
 ok1 = True
  End If
  
If (Not IsNumeric(txtprice.Text)) Then
 lblprice.Visible = True
 ok2 = False
 Else
 lblprice.Visible = False
 ok2 = True
 End If
 
 If (Not IsNumeric(txtupdateprice.Text)) Then
 lblpupdate.Visible = True
 ok3 = False
 Else
 lblpupdate.Visible = False
 ok3 = True
 End If
 
 If (Not IsNumeric(txtupdatequantity.Text)) Then
 lblsupdate.Visible = True
 ok4 = False
 Else
 lblsupdate.Visible = False
 ok4 = True
 End If
 
If (ok1 And ok2 And ok3 And ok4) = True Then
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

Private Sub txtprice_Change()
If Trim(txtprice.Text) = "" Then
    lblprice.Visible = True
    Else
    lblprice.Visible = False
End If
End Sub
Private Sub txtupdatequantity_Change()
If Trim(txtupdatequantity.Text) = "" Then
   lblsupdate.Visible = True
    Else
    lblsupdate.Visible = False
End If
End Sub
Private Sub txtupdateprice_Change()
If Trim(txtupdateprice.Text) = "" Then
     lblpupdate.Visible = True
    Else
     lblpupdate.Visible = False
End If
End Sub

Private Sub subclearlabel()
lblquantity.Visible = False
lblprice.Visible = False
lblpupdate.Visible = False
lblsupdate.Visible = False
End Sub

