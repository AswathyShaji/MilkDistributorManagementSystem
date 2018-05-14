VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_CHART 
   Caption         =   "QUALITY CHART"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10515
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
   ScaleHeight     =   5790
   ScaleWidth      =   10515
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
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   4455
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
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
         Height          =   615
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   975
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
      Height          =   5055
      Left            =   4560
      TabIndex        =   11
      Top             =   240
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid gridchart 
         Height          =   3615
         Left            =   360
         TabIndex        =   12
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
         TabIndex        =   14
         Top             =   240
         Width           =   2415
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
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.ComboBox combomtype 
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
         Left            =   2760
         TabIndex        =   1
         Text            =   ".....select........"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtcost 
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
         Left            =   2760
         TabIndex        =   3
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox txtquality 
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
         Left            =   2760
         TabIndex        =   2
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MILK TYPE"
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
         TabIndex        =   17
         Top             =   1680
         Width           =   1065
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
         Left            =   2520
         TabIndex        =   16
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label5 
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
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblquality 
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
         Left            =   3600
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   8
         Top             =   3240
         Width           =   2205
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUALITY "
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
         TabIndex        =   7
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUALITY CHART"
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
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   3090
      End
   End
End
Attribute VB_Name = "FRM_CHART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim STRSQL As String

Public Sub subinsert()
STRSQL = " INSERT INTO TBL_CHART (QUALITY,MT_NAME,QL_COST) VALUES " _
         & " ('" & txtquality.Text & "' ,'" & combomtype.List(combomtype.ListIndex) & "'," _
         & " '" & txtcost.Text & "' )"
Set RS = adocn.Execute(STRSQL)
End Sub
Private Sub SUBMTYPE()
STRSQL = "SELECT * FROM TBL_MILKTYPE "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combomtype.AddItem (RS!MT_NAME)
RS.MoveNext
Loop
End Sub


Public Sub subClear()
txtquality.Text = ""
txtcost.Text = ""
End Sub
Private Sub cmdadd_Click()
If combomtype.Text = ".....select........" Or combomtype.Text = "" Then
 MsgBox "select the milk type"
 Else
If fnValidation = True Then
subinsert
subid
            MsgBox "Success"
           subClear
           subclearlabel
           subAddToGrid
             Else
        MsgBox "Registration Failed", vbCritical
    End If
    End If
End Sub

Public Function fnValidation()
Dim ok1, ok2 As Boolean
If Trim(txtquality.Text) = "" Then
 lblquality.Visible = True
 ok1 = False
 Else
 lblquality.Visible = False
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

Private Sub Form_Load()
subAddToGrid
SUBMTYPE
subid
cmdupdate.Enabled = False
End Sub

Private Sub txtquality_Change()
If Trim(txtquality.Text) = "" Then
    lblquality.Visible = True
    Else
    lblquality.Visible = False
End If
End Sub

Private Sub txtcost_Change()
If Trim(txtcost.Text) = "" Then
    lblcost.Visible = True
    Else
    lblcost.Visible = False
End If
End Sub

Private Sub gridchart_Click()
    If gridchart.Rows > 1 Then
        cmdupdate.Enabled = True
        cmdadd.Enabled = False
        STRSQL = "select * from TBL_CHART where QL_ID = '" & gridchart.TextMatrix(gridchart.RowSel, 0) & "'"
        Set RS = adocn.Execute(STRSQL)
        lblid.Caption = RS!QL_ID
        txtquality.Text = RS!QUALITY
        txtcost.Text = RS!QL_COST
        Dim X As Integer
        For X = 0 To combomtype.ListCount - 1
        If combomtype.List(X) = RS!MT_NAME Then
        combomtype.ListIndex = X
        Exit For
        End If
        Next X
    End If
End Sub

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

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_CHART"
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
lblquality.Visible = False
lblcost.Visible = False
End Sub

Private Sub subupdate()
 STRSQL = " UPDATE TBL_CHART SET QUALITY= '" & txtquality.Text & "'," _
  & " MT_NAME= '" & combomtype.List(combomtype.ListIndex) & "'," _
  & " QL_COST='" & txtcost.Text & "' where QL_ID=' " & lblid.Caption & "'"
 
 Set RS = adocn.Execute(STRSQL)
End Sub

Private Sub cmdupdate_Click()
If fnValidation = True Then
subupdate
subid
MsgBox "Updation Succesfull"
subClear
subclearlabel
cmdupdate.Enabled = False
cmdadd.Enabled = True
subAddToGrid

           Else
            MsgBox "Updation Failed", vbCritical
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
