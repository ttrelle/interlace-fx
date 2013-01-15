VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form f__main 
   Caption         =   "CyCo Interlacer"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   Icon            =   "f__main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   10530
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton b_reopen 
      Caption         =   "Undo"
      Height          =   375
      Left            =   1260
      TabIndex        =   6
      Top             =   60
      Width           =   1095
   End
   Begin VB.Frame f_options 
      Caption         =   "Options"
      Height          =   4695
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   2295
      Begin VB.PictureBox pb_color_over 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   900
         ScaleHeight     =   525
         ScaleWidth      =   705
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox tb_gap 
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Text            =   "1"
         Top             =   300
         Width           =   495
      End
      Begin VB.ListBox lb_mode 
         Height          =   3180
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   2055
      End
      Begin VB.PictureBox pb_color 
         Appearance      =   0  '2D
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   120
         ScaleHeight     =   525
         ScaleWidth      =   705
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Gap"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton b_save 
      Caption         =   "Save BMP"
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton b_interlace 
      Caption         =   "Interlace"
      Height          =   375
      Left            =   1260
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton b_open 
      Caption         =   "Open"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1095
   End
   Begin VB.PictureBox pb_img 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7575
      Left            =   2460
      ScaleHeight     =   7515
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   540
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "f__main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim fn$

Private Sub b_interlace_Click()
    
    interlace pb_img, Val(tb_gap.Text), lb_mode.ListIndex + 1, pb_color.backcolor

End Sub

Sub openIt()

    If fn <> "" Then
    
        pb_img.Picture = LoadPicture(fn)
    
    End If

End Sub

Private Sub b_open_Click()
    
    fn = getFileName(cd, , "Bitmaps|*.bmp;*.jpg;*.gif;*.ico")
    openIt
    
End Sub


Private Sub b_reopen_Click()
    openIt
End Sub

Private Sub b_save_Click()

    Dim fn2$
    fn2 = getSaveFileName(cd, , "Bitmap (*.bmp)|*.bmp")
    If fn2 <> "" Then
    
        SavePicture pb_img.Picture, fn2
    
    End If

End Sub

Private Sub Command1_Click()

    pb_img.Picture = Me.Icon

End Sub

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
        .ListIndex = vbCopyPen - 1
         
    End With
    
End Sub

Private Sub pb_color_Click()

    With cd
    
        .ShowColor
        pb_color.backcolor = .Color
        
    End With

End Sub

Private Sub pb_img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    pb_color.backcolor = pb_img.Point(X, Y)

End Sub

Private Sub pb_img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    pb_color_over.backcolor = pb_img.Point(X, Y)

End Sub
