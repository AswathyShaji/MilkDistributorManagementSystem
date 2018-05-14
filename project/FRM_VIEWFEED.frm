VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_VIEWFEED 
   Caption         =   "FEED DETAILS"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7425
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin MSFlexGridLib.MSFlexGrid gridfeed 
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
         Caption         =   "FEED DETALIS"
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
         Width           =   3675
      End
   End
End
Attribute VB_Name = "FRM_VIEWFEED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()
subAddToGrid
End Sub
