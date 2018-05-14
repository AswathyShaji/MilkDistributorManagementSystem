VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_VIEWUSER 
   Caption         =   "USER DETAILS"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15450
   LinkTopic       =   "Form2"
   ScaleHeight     =   7485
   ScaleWidth      =   15450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin MSFlexGridLib.MSFlexGrid griduser 
         Height          =   5775
         Left            =   480
         TabIndex        =   1
         Top             =   1200
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   10186
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER DETAILS"
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
         Left            =   4680
         TabIndex        =   2
         Top             =   360
         Width           =   3660
      End
   End
End
Attribute VB_Name = "FRM_VIEWUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub subSetgrid()
    griduser.Cols = 7
    griduser.Rows = 2
    griduser.FixedRows = 1
    griduser.TextMatrix(0, 1) = "SL No"
    griduser.TextMatrix(0, 2) = "Name"
    griduser.TextMatrix(0, 3) = "Address"
    griduser.TextMatrix(0, 4) = "Phone Number"
    griduser.TextMatrix(0, 5) = "Email id"
    griduser.TextMatrix(0, 6) = "User name"
    griduser.ColWidth(0) = 0
    griduser.ColWidth(1) = 750
    griduser.ColWidth(2) = 1730
    griduser.ColWidth(3) = 4000
    griduser.ColWidth(4) = 1600
    griduser.ColWidth(5) = 2800
    griduser.ColWidth(6) = 2800
End Sub

Public Sub subAddToGrid()
    griduser.Clear
    subSetgrid
    STRSQL = "select * from TBL_USERINF"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            griduser.TextMatrix(i, 0) = RS!U_ID
            griduser.TextMatrix(i, 1) = SLNO
            griduser.TextMatrix(i, 2) = RS!U_NAME
            griduser.TextMatrix(i, 3) = RS!U_ADDRESS
            griduser.TextMatrix(i, 4) = RS!U_PH
            griduser.TextMatrix(i, 5) = RS!U_EMAIL
            griduser.TextMatrix(i, 6) = RS!U_USERNAME
            griduser.Rows = griduser.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    griduser.Rows = griduser.Rows - 1
End Sub

Private Sub Form_Load()
subAddToGrid
End Sub
