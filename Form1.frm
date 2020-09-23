VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crime Fighting Console"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   3075
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   495
         Left            =   300
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   495
         Left            =   1620
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   3
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Text1"
         BuddyDispid     =   196618
         OrigLeft        =   2040
         OrigTop         =   120
         OrigRight       =   2235
         OrigBottom      =   915
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ForeColor:"
         Height          =   195
         Left            =   720
         TabIndex        =   7
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor:"
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   900
         TabIndex        =   5
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1740
         TabIndex        =   4
         Top             =   780
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   1860
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.NRScroll NRScroll1 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   238
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Click for Knight Rider Version"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   180
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Click for SCANNER Version"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   2955
   End
   Begin VB.Label lblAbout 
      Caption         =   "Label5"
      Height          =   2115
      Left            =   60
      TabIndex        =   13
      Top             =   3300
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox "Good Morning, Michael" & vbCrLf & "Kit reporting.." & vbCrLf & "All Systems Go!"
    NRScroll1.StartScroll
End Sub

Private Sub Command2_Click()
   NRScroll1.EndScroll
   MsgBox "This is Kit - Signing off."
End Sub

Private Sub Command3_Click()
    Height = Height + 1000
    Frame1.Top = Frame1.Top + 1000
    NRScroll1.Height = NRScroll1.Height + 1000
    lblAbout.Top = lblAbout.Top + 1000
    Command3.Visible = False
    Command4.Visible = True
End Sub

Private Sub Command4_Click()
    Height = Height - 1000
    Frame1.Top = Frame1.Top - 1000
    NRScroll1.Height = NRScroll1.Height - 1000
    lblAbout.Top = lblAbout.Top - 1000
    Command4.Visible = False
    Command3.Visible = True
End Sub

Private Sub Form_Load()
    Label3(0).BackColor = NRScroll1.ForeColor
    Label3(1).BackColor = NRScroll1.BackColor
    Text1 = NRScroll1.Speed
    
    lblAbout = "First of all, this code is not 100% mine. The code for the gradient is from Kath Rock (AWESOME CODE!). Visit KR's code (No. 6154) and vote (for both of us <GRIN>)." & vbCrLf & vbCrLf & "Sparq - <jay@alphamedia.net>"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label3_Click(Index As Integer)
    On Error GoTo Err:
    With cdlg1
        .CancelError = True
        .ShowColor
        Label3(Index).BackColor = .Color
    End With
    NRScroll1.ForeColor = Label3(0).BackColor
    NRScroll1.BackColor = Label3(1).BackColor

Err:
End Sub

Private Sub Text1_Change()
    NRScroll1.Speed = Val(Text1)
End Sub
