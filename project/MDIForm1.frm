VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10650
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   20250
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7320
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu USER 
      Caption         =   "USER MANAGEMENT"
      Begin VB.Menu ADD 
         Caption         =   "REGISTRATION"
      End
      Begin VB.Menu SALARY 
         Caption         =   "SET SALARY"
      End
      Begin VB.Menu ATTENEDENCE 
         Caption         =   "ATTENDENCE MARKING"
      End
      Begin VB.Menu USERPAYMENT 
         Caption         =   "PAYMENT"
      End
   End
   Begin VB.Menu FARMER 
      Caption         =   "FARMER MANAGEMENT"
      Begin VB.Menu NEW 
         Caption         =   "REGISTRATION"
      End
      Begin VB.Menu PAYMENT 
         Caption         =   "PAYMENT"
      End
   End
   Begin VB.Menu FEED 
      Caption         =   "CATTLE FEED"
      Begin VB.Menu CATEGORY 
         Caption         =   "FEED CATEGORY"
      End
      Begin VB.Menu FEED_DETALIS 
         Caption         =   "FEED DETAILS"
      End
   End
   Begin VB.Menu MILLK 
      Caption         =   "MILK MANAGEMENT"
      Begin VB.Menu MTYPE 
         Caption         =   "MILK YPES"
      End
      Begin VB.Menu CHART 
         Caption         =   "QUALITY CHART"
      End
      Begin VB.Menu MCOLLECTION 
         Caption         =   "MILK COLLECTION"
      End
   End
   Begin VB.Menu SALES 
      Caption         =   "SALES MANAGEMENT"
      Begin VB.Menu MSALE 
         Caption         =   "MILK SALE"
      End
      Begin VB.Menu FSALE 
         Caption         =   "FEED SALE"
      End
   End
   Begin VB.Menu VIEW 
      Caption         =   "VIEW"
      Begin VB.Menu VIEWSTAFF 
         Caption         =   "STAFF DETAILS"
      End
      Begin VB.Menu ATTENDENCE 
         Caption         =   "STAFF ATTENDENCE"
      End
      Begin VB.Menu FPAYMENT 
         Caption         =   "FARMER PAYMENT"
      End
      Begin VB.Menu FCAEGORYVIEW 
         Caption         =   "FEED CATEGORIES"
      End
      Begin VB.Menu FEEDVIEW 
         Caption         =   "FEED DETAILS"
      End
      Begin VB.Menu VIWECHART 
         Caption         =   "QUALITY CHART"
      End
   End
   Begin VB.Menu report 
      Caption         =   "REPORT"
      Begin VB.Menu SALE_REPORT 
         Caption         =   "FEED SALE REPORT"
      End
      Begin VB.Menu MILKSALE 
         Caption         =   "MILK SALE REPORT"
      End
      Begin VB.Menu MSTOCKREPORT 
         Caption         =   "MILK STOCK REPORT"
      End
   End
   Begin VB.Menu SETTINGS 
      Caption         =   "SETTINGS"
      Begin VB.Menu CHANGE_PASSWORD 
         Caption         =   "CHANGE PASSWORD"
      End
      Begin VB.Menu LOGOUT 
         Caption         =   "LOGOUT"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ADD_Click()
FRM_USER.Show
End Sub

Private Sub ATTENDENCE_Click()
FRM_VIEWATTENDANCE.Show
End Sub

Private Sub ATTENDENCEVIEW_Click()
FRM_VIEWATTENDENCE.Show
End Sub

Private Sub ATTENEDENCE_Click()
FRM_ATTENDANCE.Show
End Sub

Private Sub CATEGORY_Click()
FRM_CATEGORY.Show
End Sub

Private Sub CATEGORYVIEW_Click()
FRM_VIEWCATEGORY.Show
End Sub

Private Sub CHANGE_PASSWORD_Click()
FRM_CHANGE_PASSWORD.Show
End Sub

Private Sub CHART_Click()
FRM_CHART.Show
End Sub

Private Sub CHARTVIEW_Click()
FRM_VIEWCHART.Show
End Sub

Private Sub FCAEGORYVIEW_Click()
FRM_VIEWCATEGORY.Show
End Sub

Private Sub FEED_DETALIS_Click()
FRM_FEED.Show
End Sub

Private Sub FEEDVIEW_Click()
FRM_VIEWFEED.Show
End Sub

Private Sub FPAYMENT_Click()
FRM_VIEWFARMERPAYMENT.Show
End Sub

Private Sub FSALE_Click()
FRM_FEEDSALE.Show
End Sub

Private Sub LOGOUT_Click()
Dim r As Integer
r = MsgBox("Are you sure?", vbYesNo, "WARNING")
If r = 6 Then
Unload Me
End If
End Sub

Private Sub MCOLLECTION_Click()
FRM_MCOLLECTION.Show
End Sub

Private Sub MILKSALE_Click()
CrystalReport1.ReportFileName = App.Path & "/REPORT/MILKSALE1.rpt"
      CrystalReport1.RetrieveDataFiles
      CrystalReport1.WindowState = crptMaximized
      CrystalReport1.Action = 1
End Sub

Private Sub MSALE_Click()
FRM_MILKSALE.Show
End Sub

Private Sub MSTOCK_Click()
FRM_STOCK.Show
End Sub

Private Sub MSTOCKREPORT_Click()
CrystalReport1.ReportFileName = App.Path & "/REPORT/MILKCOLLECTION.rpt"
      CrystalReport1.RetrieveDataFiles
      CrystalReport1.WindowState = crptMaximized
      CrystalReport1.Action = 1
End Sub

Private Sub MTYPE_Click()
FRM_MTYPE.Show
End Sub

Private Sub NEW_Click()
FRM_FARMER.Show
End Sub

Private Sub PAYMENT_Click()
FRM_FARMERPAYMENT.Show
End Sub

Private Sub UPDATEFEED_Click()
FRM_FEEDUPDATE.Show
End Sub

Private Sub PAYMENTTVIEW_Click()
FRM_VIEWPAYMENT.Show
End Sub

Private Sub SALARY_Click()
FRM_USERSALARY.Show
End Sub

Private Sub SALE_REPORT_Click()
      CrystalReport1.ReportFileName = App.Path & "/REPORT/SALES.rpt"
      CrystalReport1.RetrieveDataFiles
      CrystalReport1.WindowState = crptMaximized
      CrystalReport1.Action = 1
    
End Sub

Private Sub USERDATA_Click()
FRM_VIEWUSER.Show
End Sub

Private Sub USERPAYMENT_Click()
FRM_USERPAYMENT.Show
End Sub

Private Sub VIEWSTAFF_Click()
FRM_VIEWUSER.Show
End Sub

Private Sub VIWECHART_Click()
FRM_VIEWCHART.Show
End Sub
