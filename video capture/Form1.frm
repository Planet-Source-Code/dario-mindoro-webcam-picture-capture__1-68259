VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Capture Utility - Dario Mindoro"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save Picture"
      Height          =   345
      Left            =   2850
      TabIndex        =   4
      Top             =   2310
      Width           =   1185
   End
   Begin VB.CommandButton cmdVidFormat 
      Caption         =   "Video Format"
      Height          =   345
      Left            =   2880
      TabIndex        =   3
      Top             =   90
      Width           =   1185
   End
   Begin VB.CheckBox chkVideo 
      Caption         =   "Start Video"
      Height          =   345
      Left            =   2880
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   450
      Width           =   1185
   End
   Begin VB.Timer vTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2790
      Top             =   1740
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   2880
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   870
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      DrawWidth       =   3
      Height          =   2565
      Left            =   90
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   0
      Top             =   90
      Width           =   2685
      Begin VB.Shape shp1 
         BorderColor     =   &H0000FF00&
         Height          =   1350
         Left            =   750
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FOR WEBCAM DECLARATIONS
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

Private mCapHwnd As Long
Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054

Dim SELECTOR As tRECT




Private Sub chkVideo_Click()
    Dim c As Long
    
    'enable/disable video
    If chkVideo.Value = vbChecked Then
        vTimer.Enabled = True
        chkVideo.Caption = "Stop Video"
        StartCam
    Else
        vTimer.Enabled = False
        chkVideo.Caption = "Start Video"
        StopCam
    End If
End Sub

Private Sub cmdVidFormat_Click()
    'display the video format dialog for adjustment
    capDlgVideoFormat mCapHwnd
End Sub




Private Sub Command1_Click()
    Dim fname As String
    fname = Format(Now, "mmddyyhms")
    SavePicture Picture2.Image, fname & ".jpg"
    MsgBox "Image Saved"
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rc As tRECT
    If Button = 1 Then
        shp1.left = x
        shp1.top = y
    
    'copy image from the video frame
    rc.left = shp1.left
    rc.top = shp1.top
    rc.height = shp1.height
    rc.width = shp1.width
    
    'copy what on inside the green rectable to target picture box
    CopyBuffer Picture1, Picture2, rc
        
    End If
End Sub

Private Sub vTimer_Timer()
    Dim rc As tRECT

    DoEvents
    SendMessage mCapHwnd, GET_FRAME, 0, 0
    SendMessage mCapHwnd, COPY, 0, 0
    Picture1.Picture = Clipboard.GetData
    Clipboard.Clear
    
    DoEvents
    
    'copy image from the video frame
    rc.left = shp1.left
    rc.top = shp1.top
    rc.height = shp1.height
    rc.width = shp1.width
    
    'copy what on inside the green rectable to target picture box
    CopyBuffer Picture1, Picture2, rc
    
    DoEvents

End Sub
Sub StartCam()
    'start the webcam
    mCapHwnd = capCreateCaptureWindow("Picture Capture", 0, 0, 0, 320, 240, Me.hWnd, 0)
    DoEvents
    If capDriverConnect(mCapHwnd, 0) = True Then
        vTimer.Enabled = True
    Else
        MsgBox "Capture device not installed", vbOKOnly, "Capture device Error"
    End If
End Sub

Sub StopCam()
    'stop webcam
    DoEvents
    SendMessage mCapHwnd, DISCONNECT, 0, 0
End Sub
