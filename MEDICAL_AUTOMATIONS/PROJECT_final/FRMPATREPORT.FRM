VERSION 5.00
Begin VB.Form frmpatreport 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PATIENT REPORT DETAILS"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MNUREPORT 
      Caption         =   "REPORT"
      Begin VB.Menu MNUADMISSIONPATIENTREPORT 
         Caption         =   "ADMISSION PATIENT REPORT"
      End
      Begin VB.Menu MNUDISCHARGEPATIENTREPORT 
         Caption         =   "DISCHARGE PATIENT REPORT"
      End
      Begin VB.Menu MNUOPDPATIENTREPORT 
         Caption         =   "OPD PATIENT REPORT"
      End
      Begin VB.Menu MNUBILLINGREPORT 
         Caption         =   "BILLING REPORT"
      End
      Begin VB.Menu MNUEXIT 
         Caption         =   "EXIT"
      End
   End
End
Attribute VB_Name = "frmpatreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MNUADMISSIONPATIENTREPORT_Click()
RepAdmn.Show
End Sub

Private Sub MNUBILLINGREPORT_Click()
RepBill.Show
End Sub

Private Sub MNUDISCHARGEPATIENTREPORT_Click()
RepDisc.Show
End Sub

Private Sub MNUEXIT_Click()
frmmain.Show
End Sub

Private Sub MNUOPDPATIENTREPORT_Click()
RepOPD.Show
End Sub
