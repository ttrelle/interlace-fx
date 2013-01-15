VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm f_main 
   BackColor       =   &H8000000C&
   Caption         =   "CyCo Interlacer"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8775
   Icon            =   "f_main.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manuell
   StartUpPosition =   1  'Fenstermitte
   WindowState     =   2  'Maximiert
   Begin MSComDlg.CommonDialog cd 
      Left            =   8280
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnTopFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnFile 
         Caption         =   "&Open"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnFile 
         Caption         =   "&Save BMP"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnFile 
         Caption         =   "E&xit"
         Index           =   3
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnTopEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnEdit 
         Caption         =   "&Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnEdit 
         Caption         =   "&Interlace"
         Index           =   2
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnTopAbout 
      Caption         =   "&?"
   End
End
Attribute VB_Name = "f_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MN_OPEN = 0
Const MN_SAVE = 1
Const MN_EXIT = 3

Const MN_UNDO = 0
Const MN_INTERLACE = 2

Const IMGTYPES = ".gif.bmp.jpg"

Dim lacer As New f_interlacer
Dim imgDict As Object
Dim f As f_pic

Public Sub setActiveForm(fp As f_pic)

    Set f = fp

End Sub

Private Sub addImage(fn As String)

    If Trim(fn) = "" Then Exit Sub
    If Len(fn) < 4 Then Exit Sub
    
    Dim ext$
    ext = Right(fn, 4)
    If InStrRev(IMGTYPES, ext) = 0 Then Exit Sub

    If Not (imgDict.Exists(fn)) Then

        Dim nf As New f_pic
        Dim X!, Y!
        X = 0: Y = 0

        nf.setFile fn
        nf.Icon = Me.Icon
        
        If (lacer.Left = 0) And (lacer.Top = 0) Then
            
            X = X + lacer.Width
            
        Else
            
            If Not (f Is Nothing) Then
            
                X = f.Left + 300
                Y = f.Top + 300
            
            End If
            
        End If
            
        nf.Move X, Y
        nf.Show
        
        imgDict.Add fn, nf
        
        Set f = nf
        
    Else
    
        Set f = imgDict(fn)
        f.SetFocus
        
    End If

End Sub

Public Sub removeImage(fn As String)

    If imgDict.Exists(fn) Then
        
        imgDict.Remove fn
        
    End If

End Sub

Public Sub doInterlace()

    With lacer

        f.doInterlace Val(.tb_gap.Text), _
             .lb_mode.ListIndex + 1, _
             .pb_color.backcolor, _
             Val(.tb_offset.Text)

    End With

End Sub

Public Sub setPicker(X As Single, Y As Single, col As Long)

    lacer.l_pos.Caption = "x: " & X & "   y: " & Y
    lacer.pb_color_over.backcolor = col

End Sub

Public Sub setColor(col As Long)

    lacer.pb_color.backcolor = col

End Sub

Private Sub MDIForm_Load()

    Set imgDict = CreateObject("Scripting.Dictionary")

    lacer.Move 0, 0
    lacer.Show
    Set f = Nothing

End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim fn$

    If Not Data.GetFormat(vbCFFiles) Then
        Effect = 0
        fn = ""
    Else
        fn = Data.Files(1)
    End If

    addImage fn

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Unload lacer

End Sub

Private Sub mnEdit_Click(Index As Integer)

    Select Case Index
    
        Case MN_INTERLACE: doInterlace
        Case MN_UNDO: f.doUndo
        
    End Select

End Sub

Private Sub mnFile_Click(Index As Integer)


    Select Case Index
    
        Case MN_OPEN
        
            Dim fn$
            fn = getFileName(cd, , "Images|*.jpg;*.bmp;*.gif")
            If fn <> "" Then addImage fn
    
        Case MN_SAVE
        
            f.doSave cd
    
        Case MN_EXIT
        
            Unload Me
    
    End Select
    

End Sub
