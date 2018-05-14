VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_VIEWCATEGORY 
   Caption         =   "CATEGORY DETAILS"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form2"
   ScaleHeight     =   5655
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid gridcategory 
         Height          =   4095
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   7223
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORY DETAILS"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   4980
      End
   End
End
Attribute VB_Name = "FRM_VIEWCATEGORY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub subSetgrid()
    gridcategory.Cols = 3
    gridcategory.Rows = 2
    gridcategory.FixedRows = 1
    gridcategory.TextMatrix(0, 0) = "ID"
    gridcategory.TextMatrix(0, 1) = "SL No"
    gridcategory.TextMatrix(0, 2) = "Name"
    gridcategory.ColWidth(0) = 1000
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

Private Sub Form_Load()
subAddToGrid
End Sub
