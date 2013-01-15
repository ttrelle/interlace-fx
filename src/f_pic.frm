VERSION 5.00
Begin VB.Form f_pic 
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "f_pic.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5400
      Width           =   6375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5235
      Left            =   6480
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox pb_pic 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   0
      Top             =   0
      Width           =   2835
      Begin VB.PictureBox pb_src 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   1995
         Left            =   360
         MousePointer    =   2  'Kreuz
         ScaleHeight     =   133
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   3
         Top             =   300
         Width           =   2595
      End
   End
End
Attribute VB_Name = "f_pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim sizex&, sizey&
Dim fn$

Public Sub doSave(cd As CommonDialog)

    Dim fn2$
    fn2 = getFileName(cd, "Save Bitmap", "Windows BMP (*.bmp)|*.bmp")
    If fn2 <> "" Then
    
        SavePicture pb_src.Picture, fn2
    
    End If

End Sub

Public Sub doUndo()

    pb_src.Picture = LoadPicture(fn)

End Sub

Public Sub doInterlace(ByVal gap As Long, mode As Integer, col As Long, offset As Long)

    interlace pb_src, gap, mode, col, offset

End Sub

Sub updateLabel()

    With pb_src

        .Left = -HScroll1.Value
        .Top = -VScroll1.Value
        .SetFocus
        
    End With

End Sub

Public Sub setFile(filename As String)

    fn = filename

    With pb_src

        .Picture = LoadPicture(fn)
        Caption = fn + " [" & .Width & "x" & .Height & "]"
        Width = .ScaleX(.Picture.Width, vbHimetric, vbTwips) + .ScaleX(VScroll1.Width, vbPixels, vbTwips)
        Height = .ScaleY(.Picture.Height, vbHimetric, vbTwips) + .ScaleY(HScroll1.Height, vbPixels, vbTwips)
        
    End With
    
    Form_Resize
    
    updateLabel

End Sub

Private Sub Form_Activate()
    f_main.setActiveForm Me
End Sub

Private Sub Form_GotFocus()
    f_main.setActiveForm Me
End Sub

Private Sub Form_Resize()

    If WindowState = vbMinimized Then Exit Sub
    If Not Visible Then Exit Sub
    
    sizex = pb_src.Width
    sizey = pb_src.Height
    
    If ScaleWidth < sizex Then
    
        HScroll1.Visible = -1
        HScroll1.Top = ScaleHeight - HScroll1.Height
        VScroll1.Height = ScaleHeight - pb_pic.Top - HScroll1.Height
        pb_pic.Height = ScaleHeight - pb_pic.Top - HScroll1.Height
    
    Else
    
        HScroll1.Visible = 0
        VScroll1.Height = ScaleHeight - pb_pic.Top
        pb_pic.Height = ScaleHeight - pb_pic.Top
        
    
    End If

    If ScaleHeight < sizey Then
    
        VScroll1.Visible = -1
        VScroll1.Left = ScaleWidth - VScroll1.Width
        HScroll1.Width = ScaleWidth - VScroll1.Width
        pb_pic.Width = ScaleWidth - VScroll1.Width
    
    Else
    
        VScroll1.Visible = 0
        HScroll1.Width = ScaleWidth
        pb_pic.Width = ScaleWidth
        
    End If
        
    If (ScaleWidth >= sizex) And (ScaleHeight >= sizey) Then
    
        With pb_src
        
            .Left = (pb_pic.Width - .Width) / 2
            .Top = (pb_pic.Height - .Height) / 2
            
        End With
    
    End If
        
    SetScroll
    'pb_pic.Refresh

End Sub

Sub SetScroll()

    If pb_pic.ScaleWidth < sizex Then
    'If pb_pic.ScaleWidth < pb_pic.Picture.Width Then
    
        HScroll1.Min = 0
        'HScroll1.Max = pb_pic.Picture.Width - pb_pic.ScaleWidth
        HScroll1.Max = sizex - pb_pic.ScaleWidth
        HScroll1.LargeChange = pb_pic.ScaleWidth
        HScroll1.SmallChange = HScroll1.LargeChange / 10
    
    Else
    
        HScroll1.Value = 0
        
    End If

    If pb_pic.ScaleHeight < sizey Then
    'If pb_pic.ScaleHeight < pb_pic.Picture.Height Then
    
        VScroll1.Min = 0
        'VScroll1.Max = pb_pic.Picture.Height - pb_pic.ScaleHeight
        VScroll1.Max = sizey - pb_pic.ScaleHeight
        VScroll1.LargeChange = pb_pic.ScaleHeight
        VScroll1.SmallChange = VScroll1.LargeChange / 10
    
    Else
    
        VScroll1.Value = 0
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    f_main.removeImage fn

End Sub

Private Sub HScroll1_Change()
    'pb_pic.Refresh
    updateLabel
End Sub

Private Sub HScroll1_Scrol()
    'pb_pic.Refresh
    updateLabel
End Sub

Private Sub pb_src_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Button
    
        Case vbLeftButton
    
            f_main.setColor pb_src.Point(X, Y)
    
        Case vbMiddleButton
        
            doUndo
        
        Case vbRightButton
        
            f_main.doInterlace
    
    End Select


End Sub

Private Sub pb_src_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    f_main.setPicker X, Y, pb_src.Point(X, Y)

End Sub

Private Sub VScroll1_Change()
    'pb_pic.Refresh
    updateLabel

End Sub

Private Sub VScroll1_Scroll()
    'pb_pic.Refresh
    updateLabel
End Sub
