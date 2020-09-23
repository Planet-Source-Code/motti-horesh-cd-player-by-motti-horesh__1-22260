VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CD PLAYER"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CD Player"
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   720
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.Label L3abel1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Track Number:"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   840
         Picture         =   "Form1.frx":0000
         Top             =   360
         Width           =   360
      End
      Begin VB.Image Image4 
         Height          =   270
         Left            =   1560
         Picture         =   "Form1.frx":03B3
         Top             =   360
         Width           =   570
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   2280
         Picture         =   "Form1.frx":08A0
         Top             =   360
         Width           =   750
      End
      Begin VB.Image Image2 
         Height          =   285
         Left            =   240
         Picture         =   "Form1.frx":0DF4
         Top             =   360
         Width           =   330
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   3240
         Picture         =   "Form1.frx":1199
         Top             =   360
         Width           =   645
      End
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////
'//                                     //
'//     CD PLAYER BY MOTTI HORESH       //
'//    ----------------------------     //
'//                                     //
'//   I've made this program to test    //
'//   The Control "MMControl" this      //
'//   this program used the control for //
'//   3D buttom also,  have any comments//
'//   question or bugs,  e-mail me at:  //
'//   h_motti@hotmail.com               //
'//                                     //
'/////////////////////////////////////////



Private Sub Form_Load()
' call to CDAUDIO - CD Driver -
MMControl1.DeviceType = "CDAUDIO"
' Read the Data from the CD
MMControl1.Command = "open"
' Put the Track That playing
Label1.Caption = MMControl1.Track
End Sub

Private Sub Form_Terminate()
' when the program is closed the song stopped
MMControl1.Command = "stop"
End Sub

Private Sub Image1_Click()
' stop the song
MMControl1.Command = "stop"
' Display the Track Number
Label1.Caption = MMControl1.Track
End Sub

Private Sub Image2_Click()
' play the prev track
MMControl1.Command = "Prev"
' Display the track number
Label1.Caption = MMControl1.Track
End Sub

Private Sub Image3_Click()
' puase the corrent playing song, to continue the song press "play"
MMControl1.Command = "Pause"
End Sub

Private Sub Image4_Click()
Dim Now As String
' start playing the track and display the track # at label1.caption
MMControl1.Command = "Play"
Label1.Caption = MMControl1.Track
End Sub

Private Sub Image5_Click()
' plays the next track and display the track # at lable1.caption
MMControl1.Command = "Next"
Label1.Caption = MMControl1.Track
End Sub
