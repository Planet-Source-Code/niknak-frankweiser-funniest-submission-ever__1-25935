VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frankweiser"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_play 
      Caption         =   "Play"
      Height          =   555
      Left            =   5640
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.PictureBox pic_destination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5295
      Left            =   60
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      Begin VB.Timer tim_animtimer 
         Left            =   3780
         Top             =   1020
      End
      Begin MSComctlLib.ImageList ils_frankframes 
         Left            =   3120
         Top             =   1020
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   459
         ImageHeight     =   394
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":08CA
               Key             =   "look left"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":5B10
               Key             =   "drink 2"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":AE5D
               Key             =   "drink 1"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":10291
               Key             =   "drink 3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":156E7
               Key             =   "drink 4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":1AB94
               Key             =   "normal"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":2007D
               Key             =   "drink 5"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":25567
               Key             =   "look right"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":2AA54
               Key             =   "normal talk"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":2FF64
               Key             =   "drink 7"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_main.frx":35504
               Key             =   "drink 6"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox pic_buffer 
      AutoRedraw      =   -1  'True
      Height          =   5295
      Left            =   60
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   6915
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'******************************************************
'FRANKWEISER V1
'******************************************************
'THANK YOU FORM DOWNLOADING FRANKWEISER, IF YOU LIKE IT
'OR HAVE FOUND A BUG OF SOME SORT PLEASE SEND ME AN
'EMAIL TO NP24@BLUEYONDER.CO.UK.  THANKS AGAIN!
'NICK PATEMAN 2001

'******************************************************
'API DECLARATIONS
'******************************************************
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'******************************************************
'PRIVATE CONSTANTS
'******************************************************
    Private Const SND_ASYNC = &H1
    Private Const SND_NODEFAULT = &H2

'******************************************************
'PRIVATE VARIABLES
'******************************************************
    Private anim_over As Boolean
    Private scene_busy As Boolean

'******************************************************
'FRAME DRAWING ROUTINE
'DRAW INTO BUFFER THEN BLITTED TO DESTINATION
'******************************************************
    Private Sub draw_frame(ByVal Destination As PictureBox, ByVal buffer As PictureBox, frame As String, Optional stretch As Boolean)
        With buffer
            .Width = ils_frankframes.ListImages(frame).Picture.Width
            .Height = ils_frankframes.ListImages(frame).Picture.Height
            .Picture = ils_frankframes.ListImages(frame).Picture
         End With
        With Destination
            BitBlt .hdc, 0, 0, .Width, .Height, buffer.hdc, 0, 0, vbSrcCopy
            .Refresh
        End With
    End Sub

'******************************************************
'ANIMATION ROUTINES
'******************************************************
    'MAKE FRANK DRINK FOR (LENGTHMS) AMMOUNT OF TIME
    Private Sub frank_anim_drink(lengthms As Long, ByVal animtimer As Timer, Optional soundfile As String)
        Dim takebreath As Boolean
        Dim frame As Integer
        For frame = 1 To 7
            DoEvents
            draw_frame pic_destination, pic_buffer, "drink" & Str(frame)
        Next frame
        timer_reset lengthms, animtimer
        timer_start animtimer
        While Not anim_over
            DoEvents
            Sleep 100
            takebreath = Not takebreath
            If takebreath Then
                draw_frame pic_destination, pic_buffer, "drink 6"
            Else
                If soundfile <> "" Then sndPlaySound soundfile, SND_ASYNC Or SND_NODEFAULT
                draw_frame pic_destination, pic_buffer, "drink 7"
            End If
        Wend
        timer_stop animtimer
        For frame = 7 To 1 Step -1
            DoEvents
            draw_frame pic_destination, pic_buffer, "drink" & Str(frame)
        Next frame
    End Sub
    
    'MAKE FRANK LOOK LEFT FOR (LENGTHMS) AMMOUNT OF TIME
    Private Sub frank_anim_left(lengthms As Long, ByVal animtimer As Timer)
        draw_frame pic_destination, pic_buffer, "look left"
        timer_reset lengthms, animtimer
        timer_start animtimer
        While Not anim_over
            DoEvents
        Wend
        timer_stop animtimer
        draw_frame pic_destination, pic_buffer, "normal"
    End Sub

    'MAKE FRANK LOOK LEFT FOR (LENGTHMS) AMMOUNT OF TIME
    Private Sub frank_anim_right(lengthms As Long, ByVal animtimer As Timer)
        draw_frame pic_destination, pic_buffer, "look right"
        timer_reset lengthms, animtimer
        timer_start animtimer
        While Not anim_over
            DoEvents
        Wend
        timer_stop animtimer
        draw_frame pic_destination, pic_buffer, "normal"
    End Sub
    
    'MAKE FRANK MOVE HIS LIPS FOR (LENGTHMS) AMMOUNT OF TIME
    Private Sub frank_anim_talk(lengthms As Long, ByVal animtimer As Timer, Optional soundfile As String)
        Dim mouthopen As Boolean
        timer_reset lengthms, animtimer
        timer_start animtimer
        If soundfile <> "" Then sndPlaySound soundfile, SND_ASYNC Or SND_NODEFAULT
        While Not anim_over
            DoEvents
            Sleep 100
            mouthopen = Not mouthopen
            If mouthopen Then
                draw_frame pic_destination, pic_buffer, "normal talk"
            Else
                draw_frame pic_destination, pic_buffer, "normal"
            End If
        Wend
        timer_stop animtimer
        draw_frame pic_destination, pic_buffer, "normal"
    End Sub
    
    'MAKE FRANK KEEP HIS MOUTH OPEN FOR (LENGTHMS) AMMOUNT OF TIME
    Private Sub frank_anim_gawp(lengthms As Long, ByVal animtimer As Timer, Optional soundfile As String)
        Dim mouthopen As Boolean
        timer_reset lengthms, animtimer
        timer_start animtimer
        If soundfile <> "" Then sndPlaySound soundfile, SND_ASYNC Or SND_NODEFAULT
        draw_frame pic_destination, pic_buffer, "normal talk"
        While Not anim_over
            DoEvents
        Wend
        timer_stop animtimer
        draw_frame pic_destination, pic_buffer, "normal"
    End Sub
    

'******************************************************
'TIMER ROUTINES
'******************************************************
    'RESET TIMER
    Private Sub timer_reset(lengthms As Long, ByVal animtimer As Timer)
        anim_over = False
        animtimer.Interval = lengthms
    End Sub
    
    'STOP TIMER
    Private Sub timer_stop(ByVal animtimer As Timer)
        animtimer.Enabled = False
    End Sub
    
    'START TIMER
    Private Sub timer_start(ByVal animtimer As Timer)
        animtimer.Enabled = True
    End Sub

'******************************************************
'DESTINATION EVENTS
'******************************************************
    Private Sub pic_destination_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Not scene_busy Then
            If X < pic_destination.ScaleWidth / 2 - 50 Then
                draw_frame pic_destination, pic_buffer, "look left"
            ElseIf X > pic_destination.ScaleWidth / 2 + 50 Then
                draw_frame pic_destination, pic_buffer, "look right"
            Else
                draw_frame pic_destination, pic_buffer, "normal"
            End If
        End If
    End Sub

'******************************************************
'TIMER WIDGET EVENTS
'******************************************************
    Private Sub tim_animtimer_Timer()
        anim_over = True
    End Sub

'******************************************************
'FORM EVENTS
'******************************************************
    'LOAD
    Private Sub Form_Load()
        draw_frame pic_destination, pic_buffer, "normal"
    End Sub
    
    'UNLOAD
    Private Sub Form_Unload(Cancel As Integer)
        End
    End Sub

'******************************************************
'TEST AREA
'******************************************************
    Private Sub frank_anim_scene1()
        scene_busy = True
        cmd_play.Enabled = False
            frank_anim_left 2000, tim_animtimer
            frank_anim_right 2000, tim_animtimer
            frank_anim_talk 12500, tim_animtimer, App.Path & "\Sounds\obey me.wav"
            frank_anim_drink 4000, tim_animtimer, App.Path & "\Sounds\slurp.wav"
            frank_anim_left 2000, tim_animtimer
            frank_anim_right 2000, tim_animtimer
            frank_anim_gawp 1000, tim_animtimer, App.Path & "\Sounds\belch.wav"
            frank_anim_left 2000, tim_animtimer
            frank_anim_right 2000, tim_animtimer
            frank_anim_talk 2500, tim_animtimer, App.Path & "\Sounds\stink on shit.wav"
            frank_anim_drink 4000, tim_animtimer, App.Path & "\Sounds\slurp.wav"
            frank_anim_left 2000, tim_animtimer
            frank_anim_right 2000, tim_animtimer
            frank_anim_gawp 1000, tim_animtimer, App.Path & "\Sounds\belch.wav"
            frank_anim_left 2000, tim_animtimer
            frank_anim_right 2000, tim_animtimer
            frank_anim_talk 2000, tim_animtimer, App.Path & "\Sounds\evil laugh.wav"
        cmd_play.Enabled = True
        scene_busy = False
    End Sub

    Private Sub cmd_play_Click()
        If Not scene_busy Then frank_anim_scene1
    End Sub

