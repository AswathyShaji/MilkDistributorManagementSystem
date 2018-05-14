VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_FEED 
   Caption         =   "FEED DETAILS"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   11310
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
      Height          =   5895
      Left            =   5160
      TabIndex        =   17
      Top             =   480
      Width           =   5775
      Begin MSFlexGridLib.MSFlexGrid gridfeed 
         Height          =   3855
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   4695
         _ExtentX        =   8281
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
         TabIndex        =   21
         Top             =   480
         Width           =   2670
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
      Height          =   975
      Left            =   240
      TabIndex        =   16
      Top             =   5280
      Width           =   5055
      Begin VB.CommandButton cmdcancel 
         Caption         =   "CANCEL"
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
         Left            =   3120
         TabIndex        =   6
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
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   975
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
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      Begin VB.ComboBox combocategory 
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
         Left            =   3360
         TabIndex        =   1
         Text            =   ".......select................."
         Top             =   1680
         Width           =   1455
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
         Left            =   3360
         TabIndex        =   4
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox txtcname 
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
         Left            =   3360
         TabIndex        =   2
         Top             =   2400
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
         Left            =   3360
         TabIndex        =   3
         Top             =   3120
         Width           =   1455
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
         Left            =   3480
         TabIndex        =   20
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label lblcategory 
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
         Left            =   4200
         TabIndex        =   15
         Top             =   1440
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
         Left            =   4200
         TabIndex        =   14
         Top             =   3600
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
         Left            =   4080
         TabIndex        =   13
         Top             =   2880
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblcname 
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
         Left            =   4200
         TabIndex        =   12
         Top             =   2160
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRICE OF 1KG"
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
         TabIndex        =   11
         Top             =   3840
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL QUANTITY"
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
         TabIndex        =   10
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME OF CATTLE FEED"
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
         TabIndex        =   9
         Top             =   2400
         Width           =   2355
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORY OF CATTLE FEED"
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
         Top             =   1680
         Width           =   2925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATTLE FEED DETAILS"
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
         TabIndex        =   7
         Top             =   120
         Width           =   4155
      End
   End
End
Attribute VB_Name = "FRM_FEED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New Recordset
Dim STRSQL As String
Dim autIn As Integer

Private Sub SUBFEEDCATEGORY()
STRSQL = "SELECT * FROM TBL_FEEDCATEGORY "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combocategory.AddItem (RS!C_NAME)
RS.MoveNext
Loop
End Sub

Private Sub cmdadd_Click()
If combocategory.Text = ".......select................." Or combocategory.Text = "" Then
 MsgBox "select the category"
 Else
If fnValidation = True Then
subinsert
  MsgBox "Success"
    subClear
    subclearlabel
    subAddToGrid
           Else
        MsgBox "Registration Failed", vbCritical
    End If
    End If
End Sub

Private Sub Form_Load()
SUBFEEDCATEGORY
subAddToGrid
subid
End Sub
Public Sub subinsert()

STRSQL = " INSERT INTO TBL_FEED (CF_QUANTITY,CF_NAME,CF_PRICE,C_NAME) " _
          & " VALUES ('" & txtquantity.Text & "' , '" & txtcname.Text & "' , '" & txtprice.Text & "' , " _
          & " '" & combocategory.List(combocategory.ListIndex) & "')"

Set RS = adocn.Execute(STRSQL)

End Sub

Public Sub subClear()
txtquantity.Text = ""
txtcname.Text = ""
txtprice.Text = ""
End Sub

Public Function fnValidation()
Dim ok1, ok2, ok3 As Boolean
If Trim(txtcname.Text) = "" Then
 lblcname.Visible = True
 ok1 = False
 Else
 lblcname.Visible = False
 ok1 = True
 End If
 
If (Not IsNumeric(txtquantity.Text)) Then
 lblquantity.Visible = True
 ok2 = False
 Else
 lblquantity.Visible = False
 ok2 = True
  End If
  
If (Not IsNumeric(txtprice.Text)) Then
 lblprice.Visible = True
 ok3 = False
 Else
 lblprice.Visible = False
 ok3 = True
 End If
 
If (ok1 And ok2 And ok3) = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function

Private Sub txtcname_Change()
If Trim(txtcname.Text) = "" Then
    lblcname.Visible = True
    Else
    lblcname.Visible = False
End If
    
End Sub

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

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_FEED"
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
lblcname.Visible = False
lblquantity.Visible = False
lblprice.Visible = False
End Sub
Private Sub combocategory_KeyPress(KeyAscii As Integer)
a = KeyAscii
If (Not a < 8) And (Not a > 127) Then
MsgBox "Only selection allowed", vbCritical
KeyAscii = 0
Exit Sub
End If
End Sub
