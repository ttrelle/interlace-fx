Attribute VB_Name = "m_interlace"
Option Explicit
Option Base 0

Public Sub interlace(pb_img As PictureBox, gap As Long, mode As Integer, _
    backcolor As Long, offset As Long)

    Dim li&
    Dim tpp%
    tpp = Screen.TwipsPerPixelY
    
    With pb_img
    
        .AutoRedraw = True
    
        .DrawMode = mode
        .ForeColor = backcolor
    
        'li = tpp
        li = IIf(offset >= 0, offset, 0)
        While li <= .Height
        
            pb_img.Line (0, li)-(.Width, li)
            'li = li + tpp * (gap + 1)
            li = li + (gap + 1)
        
        Wend
        .Picture = .Image
        
        .AutoRedraw = False
        
    End With

End Sub

