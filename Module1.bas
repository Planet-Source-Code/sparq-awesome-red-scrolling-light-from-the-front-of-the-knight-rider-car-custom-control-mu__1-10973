Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                        ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
                        ByVal nHeight As Long, ByVal hSrcDC As Long, _
                        ByVal xSrc As Long, _
                        ByVal ySrc As Long, ByVal dwRop As Long) As Long
Global Const SRCCOPY = &HCC0020

