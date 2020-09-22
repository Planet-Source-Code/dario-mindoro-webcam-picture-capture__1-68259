Attribute VB_Name = "modGraphic"
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Type tRECT
    left As Integer
    top As Integer
    width As Integer
    height As Integer
End Type

Public rct As tRECT

'copy picture box content to another picture box
Public Sub CopyBuffer(src As PictureBox, disc As PictureBox, rct As tRECT)
    disc.Cls
    StretchBlt disc.hdc, 0, 0, disc.ScaleWidth, disc.ScaleHeight, src.hdc, rct.left + 1, rct.top + 1, rct.width - 2, rct.height - 2, vbSrcCopy
End Sub
'copy picture box content to another picture box
Public Sub DisplayPreview(srcpic As PictureBox, dispic As PictureBox, rct As tRECT)
    StretchBlt dispic.hdc, 0, 0, rct.width, rct.height, srcpic.hdc, rct.left, rct.top, rct.width, rct.height, vbSrcCopy
End Sub

