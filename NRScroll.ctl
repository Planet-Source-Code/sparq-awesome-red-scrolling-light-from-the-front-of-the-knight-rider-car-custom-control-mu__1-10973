VERSION 5.00
Begin VB.UserControl NRScroll 
   ClientHeight    =   105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ScaleHeight     =   105
   ScaleWidth      =   3900
   Begin VB.PictureBox picNR 
      BackColor       =   &H00000000&
      Height          =   120
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   4980
      TabIndex        =   0
      Top             =   0
      Width           =   5040
   End
   Begin VB.PictureBox picRight 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   330
      Left            =   2520
      ScaleHeight     =   270
      ScaleWidth      =   2790
      TabIndex        =   1
      Top             =   1980
      Width           =   2850
   End
   Begin VB.PictureBox PicLeft 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   330
      Left            =   2520
      ScaleHeight     =   270
      ScaleWidth      =   2790
      TabIndex        =   2
      Top             =   1620
      Width           =   2850
   End
End
Attribute VB_Name = "NRScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Dim mSpeed As Integer
Dim mForeColor As OLE_COLOR
Dim mBackColor As OLE_COLOR

Dim Scroll As Boolean
Public mGradient As New clsGradient

Private Sub UserControl_InitProperties()
    Scroll = False
    mBackColor = vbBlack
    mForeColor = vbRed
End Sub

Private Sub UserControl_Resize()
    With picNR
        .Top = 0
        .Left = 0
        .Width = Width
        .Height = Height
    End With
    With PicLeft
        .Top = 0
        .Left = 0
        .Width = Width / 6
        .Height = Height
    End With
    With picRight
        .Top = 0
        .Left = PicLeft.Left + PicLeft.Width + 60
        .Width = Width / 6
        .Height = Height
    End With
    
    DrawGrad
End Sub

Sub DrawGrad()
    With mGradient
        .Angle = 0
        .Color1 = mBackColor
        .Color2 = mForeColor
        .Draw UserControl.PicLeft
    End With
    PicLeft.Refresh
    
    With mGradient
        .Angle = 180
        .Color1 = mBackColor
        .Color2 = mForeColor
        .Draw UserControl.picRight
    End With
    picRight.Refresh
End Sub


Public Function EndScroll()
    Scroll = False
End Function

Public Function StartScroll()
  Dim LeftSpot As Integer
  Dim X As Integer
    LeftSpot = 0
    Scroll = True
10
    LeftSpot = -(PicLeft.Width) / 17
    Do Until LeftSpot >= (picNR.Width + PicLeft.Width) / 17
        picNR.Cls
        Call BitBlt(picNR.hDC, LeftSpot, 0, PicLeft.Width, PicLeft.Height, PicLeft.hDC, 0, 0, SRCCOPY)
        LeftSpot = LeftSpot + (1 * mSpeed)
        X = 0
        Do Until X = 1000
            DoEvents
            X = X + 1
            If Scroll = False Then GoTo Off
        Loop
    Loop
    LeftSpot = ((PicLeft.Width) / 17) * 7
    Do Until LeftSpot <= -25
        picNR.Cls
        Call BitBlt(picNR.hDC, LeftSpot, 0, PicLeft.Width, PicLeft.Height, picRight.hDC, 0, 0, SRCCOPY)
        LeftSpot = LeftSpot - (1 * mSpeed)
        X = 0
        Do Until X = 1000
            DoEvents
            X = X + 1
            If Scroll = False Then GoTo Off
        Loop
    Loop
    GoTo 10

Off:
    picNR.Cls
End Function


Public Property Get Speed() As Integer
    Speed = mSpeed
End Property

Public Property Let Speed(ByVal New_Speed As Integer)
    mSpeed = New_Speed
    PropertyChanged "Speed"
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal New_BG As OLE_COLOR)
    mBackColor = New_BG
    picNR.BackColor = New_BG
    PropertyChanged "BackColor"
    DrawGrad
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal New_FG As OLE_COLOR)
    mForeColor = New_FG
    PropertyChanged "ForeColor"
    DrawGrad
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mForeColor = PropBag.ReadProperty("ForeColor", vbRed)
    mBackColor = PropBag.ReadProperty("BackColor", vbBlack)
    picNR.BackColor = mBackColor
    mSpeed = PropBag.ReadProperty("Speed", 3)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", mForeColor, vbRed)
    Call PropBag.WriteProperty("BackColor", mBackColor, vbBlack)
    Call PropBag.WriteProperty("Speed", mSpeed, 3)
End Sub
