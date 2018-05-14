VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_FARMER 
   Caption         =   "FARMER INFORMATION"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13005
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
   ScaleHeight     =   7980
   ScaleWidth      =   13005
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
      Height          =   7215
      Left            =   6000
      TabIndex        =   21
      Top             =   480
      Width           =   6735
      Begin MSFlexGridLib.MSFlexGrid gridfarmer 
         Height          =   5415
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   9551
         _Version        =   393216
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW DETAILS"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   720
         TabIndex        =   25
         Top             =   240
         Width           =   3300
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
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   6840
      Width           =   5775
      Begin VB.CommandButton cmdupdate 
         Caption         =   "UPDATE"
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
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "CANCEL"
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
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
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
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5775
      Begin VB.ComboBox combostatus 
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
         ItemData        =   "FRM_FARMER.frx":0000
         Left            =   2160
         List            =   "FRM_FARMER.frx":000D
         TabIndex        =   5
         Text            =   "..............select................."
         Top             =   5640
         Width           =   2295
      End
      Begin VB.ComboBox combomtype 
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
         ItemData        =   "FRM_FARMER.frx":0029
         Left            =   2160
         List            =   "FRM_FARMER.frx":002B
         TabIndex        =   4
         Text            =   "..............select................."
         Top             =   4920
         Width           =   2295
      End
      Begin VB.TextBox txtph 
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
         Left            =   2175
         TabIndex        =   3
         Top             =   4200
         Width           =   2280
      End
      Begin VB.TextBox txtaddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   2055
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2520
         Width           =   2280
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2055
         TabIndex        =   1
         Top             =   1920
         Width           =   2280
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
         Left            =   2160
         TabIndex        =   23
         Top             =   1200
         Width           =   2100
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
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label lblmilktype 
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
         Left            =   3960
         TabIndex        =   19
         Top             =   4680
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblstatus 
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
         Left            =   3960
         TabIndex        =   18
         Top             =   5400
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblph 
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
         Left            =   3960
         TabIndex        =   17
         Top             =   3840
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lbladdress 
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
         Left            =   3840
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblname 
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
         Left            =   3840
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SECTION STATUS"
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
         TabIndex        =   14
         Top             =   5760
         Width           =   1740
      End
      Begin VB.Label Label6 
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
         TabIndex        =   13
         Top             =   5040
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PHONE NUMBER"
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
         TabIndex        =   12
         Top             =   4320
         Width           =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2655
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1965
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FARMER REGISTRATION"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   5565
      End
   End
End
Attribute VB_Name = "FRM_FARMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New Recordset
Dim STRSQL As String
Dim autIn As Integer

Private Sub SUBMTYPE()
STRSQL = "SELECT * FROM TBL_MILKTYPE "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
combomtype.AddItem (RS!MT_NAME)
RS.MoveNext
Loop
End Sub

Private Sub cmdadd_Click()
If combomtype.Text = "..............select................." Or combomtype.Text = "" Then
 MsgBox "select the milk type"
 Else
If combostatus.Text = "..............select................." Or combostatus.Text = "" Then
MsgBox "select the section status"
Else
If fnValidation And fnMobileValidation Then
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
  End If
           
End Sub

Private Sub cmdcancel_Click()
subClear
End Sub

Private Sub Form_Load()
SUBMTYPE
subAddToGrid
subid
cmdupdate.Enabled = False
End Sub
Public Sub subinsert()

STRSQL = " INSERT INTO TBL_FARMERINF (F_NAME,F_ADDRESS,F_PH,F_STATUS,MT_NAME,F_SECTIONSTATUS) " _
          & " VALUES ('" & txtname.Text & "' , '" & txtaddress.Text & "' , '" & txtph.Text & "' , " _
          & " '1','" & combomtype.List(combomtype.ListIndex) & "','" & combostatus.List(combostatus.ListIndex) & "')"

Set RS = adocn.Execute(STRSQL)

End Sub

Public Sub subClear()
txtname.Text = ""
txtaddress.Text = ""
txtph.Text = ""
End Sub
Public Function fnValidation()
Dim ok1, ok2, ok3 As Boolean
If Trim(txtname.Text) = "" Then
 lblname.Visible = True
 ok1 = False
 Else
 lblname.Visible = False
 ok1 = True
 End If
 
If Trim(txtaddress.Text) = "" Then
 lbladdress.Visible = True
 ok2 = False
 Else
 lbladdress.Visible = False
 ok2 = True
  End If
  
If Trim(txtph.Text) = "" Then
 lblph.Visible = True
 ok3 = False
 Else
 lblph.Visible = False
 ok3 = True
 End If
 
If (ok1 And ok2 And ok3) = True Then
fnValidation = True
Else
fnValidation = False
End If
End Function
Private Sub txtname_Change()
If Trim(txtname.Text) = "" Then
    lblname.Visible = True
    Else
    lblname.Visible = False
End If
    
End Sub

Private Sub txtaddress_Change()
If Trim(txtaddress.Text) = "" Then
    lbladdress.Visible = True
    Else
    lbladdress.Visible = False
End If
End Sub

Private Sub txtph_Change()
If Trim(txtph.Text) = "" Then
    lblph.Visible = True
    Else
    lblph.Visible = False
End If
End Sub

Public Function fnMobileValidation()
        Dim Mobile As String
        Mobile = txtph.Text
        Dim ok As Boolean
        If (IsNumeric(Mobile) And Len(Mobile) = "10") Then
            ok = True
            lblph.Visible = False
        Else
            ok = False
            lblph.Caption = "* Invalid Mobile Number"
            lblph.Visible = True
        End If
        fnMobileValidation = ok
End Function

Private Sub gridfarmer_Click()
    If gridfarmer.Rows > 1 Then
        cmdupdate.Enabled = True
        cmdadd.Enabled = False
        STRSQL = "select * from TBL_FARMERINF where F_ID = '" & gridfarmer.TextMatrix(gridfarmer.RowSel, 0) & "'"
        Set RS = adocn.Execute(STRSQL)
        lblid.Caption = RS!F_ID
        txtname.Text = RS!F_NAME
        txtaddress.Text = RS!F_ADDRESS
        txtph.Text = RS!F_PH
        Dim X As Integer
        For X = 0 To combomtype.ListCount - 1
        If combomtype.List(X) = RS!MT_NAME Then
        combomtype.ListIndex = X
        Exit For
        End If
        Next X
        Dim s As Integer
        For s = 0 To combostatus.ListCount - 1
        If combostatus.List(s) = RS!F_SECTIONSTATUS Then
        combostatus.ListIndex = s
        Exit For
        End If
        Next s
        
        
    End If
End Sub

Private Sub subSetgrid()
    gridfarmer.Cols = 5
    gridfarmer.Rows = 2
    gridfarmer.FixedRows = 1
    gridfarmer.TextMatrix(0, 1) = "SL No"
    gridfarmer.TextMatrix(0, 2) = "Name"
    gridfarmer.TextMatrix(0, 3) = "Address"
    gridfarmer.TextMatrix(0, 4) = "Phone Number"
    gridfarmer.ColWidth(0) = 0
    gridfarmer.ColWidth(1) = 750
    gridfarmer.ColWidth(2) = 1730
    gridfarmer.ColWidth(3) = 1730
    gridfarmer.ColWidth(4) = 1730
End Sub

Public Sub subAddToGrid()
    gridfarmer.Clear
    subSetgrid
    STRSQL = "select * from TBL_FARMERINF"
    Set RS = adocn.Execute(STRSQL)
    If RS.RecordCount > 0 Then
        i = 1
        SLNO = 1
        While Not RS.EOF
            gridfarmer.TextMatrix(i, 0) = RS!F_ID
            gridfarmer.TextMatrix(i, 1) = SLNO
            gridfarmer.TextMatrix(i, 2) = RS!F_NAME
            gridfarmer.TextMatrix(i, 3) = RS!F_ADDRESS
            gridfarmer.TextMatrix(i, 4) = RS!F_PH
            gridfarmer.Rows = gridfarmer.Rows + 1
            SLNO = SLNO + 1
            RS.MoveNext
            i = i + 1
        Wend
    End If
    gridfarmer.Rows = gridfarmer.Rows - 1
End Sub

Private Sub subid()
Dim RS As New ADODB.Recordset
Dim STRSQL As String
    STRSQL = "select * from TBL_FARMERINF"
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
lblname.Visible = False
lbladdress.Visible = False
lblph.Visible = False
End Sub

Private Sub subupdate()
 STRSQL = " UPDATE TBL_FARMERINF SET F_NAME= '" & txtname.Text & "'," _
 & " F_ADDRESS='" & txtaddress.Text & "',F_PH= '" & txtph.Text & "', " _
 & " MT_NAME= '" & combomtype.List(combomtype.ListIndex) & "'," _
 & " F_SECTIONSTATUS='" & combostatus.List(combostatus.ListIndex) & "' where F_ID='" & lblid.Caption & "'"
 Set RS = adocn.Execute(STRSQL)
End Sub

Private Sub subdelete()
STRSQL = " DELETE FROM TBL_FARMERINF WHERE F_ID='" & lblid.Caption & "' "
Set RS = adocn.Execute(STRSQL)
End Sub


Private Sub cmddelete_Click()
subdelete
MsgBox "deleted"
subClear
subclearlabel
subAddToGrid

End Sub

Private Sub cmdupdate_Click()
If fnValidation And fnMobileValidation Then
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
