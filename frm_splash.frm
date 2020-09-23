VERSION 5.00
Begin VB.Form frm_splash 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Frankweiser"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3915
   Icon            =   "frm_splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image img_splash 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Left            =   60
      Picture         =   "frm_splash.frx":08CA
      Top             =   60
      Width           =   3810
   End
End
Attribute VB_Name = "frm_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'FORM EVENTS
'******************************************************
    Private Sub Form_Unload(Cancel As Integer)
        Load frm_main
        frm_main.Show
    End Sub

'******************************************************
'SPLASH IMAGE EVENTS
'******************************************************
    Private Sub img_splash_Click()
        Unload Me
    End Sub
