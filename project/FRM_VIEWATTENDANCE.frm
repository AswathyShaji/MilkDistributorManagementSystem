VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_VIEWATTENDANCE 
   Caption         =   "ATTENDENCE"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.ComboBox combouid 
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
         Left            =   2040
         TabIndex        =   4
         Text            =   ".....select........"
         Top             =   1080
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid gridattendence 
         Height          =   7575
         Left            =   240
         TabIndex        =   1
         Top             =   1560
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   13361
         _Version        =   393216
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER ID"
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
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW ATTENDENCE"
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
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   5025
      End
   End
End
Attribute VB_Name = "FRM_VIEWATTENDANCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SUBUID()
STRSQL = "SELECT * FROM TBL_USERINF "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combouid.AddItem (RS!U_ID)
RS.MoveNext
Loop
End Sub

Private Sub subSetgrid()
    gridattendence.Cols = 4
    gridattendence.Rows = 2
    gridattendence.FixedRows = 1
    gridattendence.TextMatrix(0, 0) = "ID"
    gridattendence.TextMatrix(0, 1) = "SL No"
   ' gridattendence.TextMatrix(0, 2) = "USER ID"
    gridattendence.TextMatrix(0, 2) = "DATE"
    gridattendence.TextMatrix(0, 3) = "ATTENDENCE"
    gridattendence.ColWidth(0) = 0
    gridattendence.ColWidth(1) = 750
    gridattendence.ColWidth(2) = 1730
    gridattendence.ColWidth(3) = 1730
    'gridattendence.ColWidth(4) = 1730
End Sub

Public Sub subAddToGrid()
    gridattendence.Clear
    subSetgrid
    STRSQL = "select * from TBL_ATTENDANCE WHERE U_ID='" & combouid.List(combouid.ListIndex) & "'"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridattendence.TextMatrix(i, 0) = RS!A_ID
            gridattendence.TextMatrix(i, 1) = SLNO
            'gridattendence.TextMatrix(i, 2) = RS!U_ID
            gridattendence.TextMatrix(i, 2) = RS!A_DATE
            gridattendence.TextMatrix(i, 3) = RS!A_STATUS
            gridattendence.Rows = gridattendence.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridattendence.Rows = gridattendence.Rows - 1
End Sub

Private Sub combouid_Click()
subAddToGrid
End Sub

Private Sub Form_Load()
SUBUID
End Sub

