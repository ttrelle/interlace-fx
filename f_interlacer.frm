VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form f_interlacer 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Interlace"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1785
   ControlBox      =   0   'False
   Icon            =   "f_interlacer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cd 
      Left            =   1200
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame f_options 
      Caption         =   "Options"
      Height          =   4995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1755
      Begin VB.TextBox tb_offset 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   "1"
         Top             =   540
         Width           =   495
      End
      Begin VB.PictureBox pb_color_over 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   900
         ScaleHeight     =   525
         ScaleWidth      =   705
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox tb_gap 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "1"
         Top             =   180
         Width           =   495
      End
      Begin VB.ListBox lb_mode 
         Height          =   3180
         Left            =   120
         TabIndex        =   2
         Top             =   1740
         Width           =   1575
      End
      Begin VB.PictureBox pb_color 
         Appearance      =   0  '2D
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   120
         ScaleHeight     =   525
         ScaleWidth      =   705
         TabIndex        =   1
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Offset"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.Label l_pos 
         Caption         =   "x:    y:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Gap"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   435
      End
   End
End
Attribute VB_Name = "f_interlacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Sub Form_Load()

    With lb_mode
    
        .Clear
        
        .AddItem "1 - vbBlackness"
        .AddItem "2 - vbNotMergePen"
        .AddItem "3 - vbMaskNotPen"
        .AddItem "4 - vbNotCopyPen"
        .AddItem "5 - vbMaskPenNot"
        .AddItem "6 - vbInvert"
        .AddItem "7 - vbXorPen"
        .AddItem "8 - vbNotMaskPen"
        .AddItem "9 - vbMaskPen"
        .AddItem "10 - vbNotXorPen"
        .AddItem "11 - vbNop"
        .AddItem "12 - vbMergeNotPen"
        .AddItem "13 - vbCopyPen"
        .AddItem "14 - vbMergePenNot"
        .AddItem "15 - vbMergePen"
        .AddItem "16 - vbWhiteness"
        .ListIndex = 12
    
    End With

End Sub

Private Sub pb_color_Click()

    With cd
    
        .ShowColor
        pb_color.backcolor = .Color
        
    End With

End Sub

Private Sub Text1_Change()

End Sub
