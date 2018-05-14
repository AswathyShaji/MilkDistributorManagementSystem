VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_VIEWFARMERPAYMENT 
   Caption         =   "PAYMENT DETAILS"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form2"
   ScaleHeight     =   6975
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin MSFlexGridLib.MSFlexGrid gridpayment 
         Height          =   5775
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   10186
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAYMENT DETAILS"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   4920
      End
   End
End
Attribute VB_Name = "FRM_VIEWFARMERPAYMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub subSetgrid()
    gridpayment.Cols = 9
    gridpayment.Rows = 2
    gridpayment.FixedRows = 1
    gridpayment.TextMatrix(0, 1) = "SL No"
    gridpayment.TextMatrix(0, 2) = "FARMER ID"
    gridpayment.TextMatrix(0, 3) = "FROM DATE"
    gridpayment.TextMatrix(0, 4) = "TO DATE"
    gridpayment.TextMatrix(0, 5) = "MILK PAYMENT"
    gridpayment.TextMatrix(0, 6) = "FEED PRICE"
    gridpayment.TextMatrix(0, 7) = "TOTAL PAYMENT"
    gridpayment.TextMatrix(0, 8) = "BALANCE"
    gridpayment.ColWidth(0) = 0
    gridpayment.ColWidth(1) = 750
    gridpayment.ColWidth(2) = 1000
    gridpayment.ColWidth(3) = 1000
    gridpayment.ColWidth(4) = 1000
    gridpayment.ColWidth(5) = 1500
    gridpayment.ColWidth(6) = 1500
    gridpayment.ColWidth(7) = 1500
    gridpayment.ColWidth(8) = 1500
End Sub

Public Sub subAddToGrid()
    gridpayment.Clear
    subSetgrid
    STRSQL = "select * from TBL_PAYMENT"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridpayment.TextMatrix(i, 0) = RS!P_ID
            gridpayment.TextMatrix(i, 1) = SLNO
            gridpayment.TextMatrix(i, 2) = RS!F_ID
            gridpayment.TextMatrix(i, 3) = RS!FROMDATE
            gridpayment.TextMatrix(i, 4) = RS!TODATE
            gridpayment.TextMatrix(i, 5) = RS!P_MILK
            gridpayment.TextMatrix(i, 6) = RS!P_FEED
            gridpayment.TextMatrix(i, 7) = RS!P_PAYMENT
            gridpayment.TextMatrix(i, 8) = RS!P_BALANCE
            gridpayment.Rows = gridpayment.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridpayment.Rows = gridpayment.Rows - 1
End Sub

Private Sub Form_Load()
subAddToGrid
End Sub

