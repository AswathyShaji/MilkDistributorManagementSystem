VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_VIEWCHART 
   Caption         =   "QUALITY CHART"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Height          =   4695
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid gridchart 
         Height          =   3615
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   6376
         _Version        =   393216
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW CHART"
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
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FRM_VIEWCHART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub subSetgrid()
    gridchart.Cols = 5
    gridchart.Rows = 2
    gridchart.FixedRows = 1
    gridchart.TextMatrix(0, 1) = "SL No"
    gridchart.TextMatrix(0, 2) = "Quality"
    gridchart.TextMatrix(0, 3) = "Milk Type"
    gridchart.TextMatrix(0, 4) = "Cost"
    gridchart.ColWidth(0) = 0
    gridchart.ColWidth(1) = 750
    gridchart.ColWidth(2) = 1730
    gridchart.ColWidth(3) = 750
    gridchart.ColWidth(4) = 750
End Sub

Public Sub subAddToGrid()
    gridchart.Clear
    subSetgrid
    STRSQL = "select * from TBL_CHART"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridchart.TextMatrix(i, 0) = RS!QL_ID
            gridchart.TextMatrix(i, 1) = SLNO
            gridchart.TextMatrix(i, 2) = RS!QUALITY
            gridchart.TextMatrix(i, 3) = RS!MT_NAME
            gridchart.TextMatrix(i, 4) = RS!QL_COST
            gridchart.Rows = gridchart.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridchart.Rows = gridchart.Rows - 1
End Sub

Private Sub Form_Load()
subAddToGrid
End Sub
