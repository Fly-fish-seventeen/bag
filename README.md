# bag
###有格式
1.会出现控件数组不存在的问题
2.百分百，切换第二个窗体会最小化
form1代码
Private Sub Command1_Click()
    Unload Form1
    Form2.Show
    
End Sub
form2 代码
Private Sub Form_Load()
    Label1.FontSize = 30
End Sub
Public s As Integer
Public pi As Integer
Private Sub Form_Load()
    Dim d(5) As Integer
    Dim a As String
    Dim i As Integer, u As Integer, Y As Integer, t As Integer, g As Integer, j As Integer
    
    
    Timer1.Interval = 3500 '时间1毫秒秒运行一次
    Timer1.Enabled = True  '时间开始运行
    
    Randomize
    g = Int(Rnd * (4 + 1)) + 0
    d(0) = g
    Do While (1 < 2)
    Randomize
    g = Int(Rnd * (4 + 1)) + 0
    If (g <> d(0)) Then
    
    d(1) = g
    Exit Do
    End If
    Loop
    
    Do While (1 < 2)
    Randomize
    g = Int(Rnd * (4 + 1)) + 0
    If (g <> d(0) And g <> d(1)) Then
    
    d(2) = g
    Exit Do
    End If
    Loop
    
    Do While (1 < 2)
    Randomize
    g = Int(Rnd * (4 + 1)) + 0
    If (g <> d(0) And g <> d(1) And g <> d(2)) Then
    
    
    d(3) = g
    Exit Do
    End If
    Loop
    
    
    
    
    For i = 0 To 4
    If (i <> d(0) And i <> d(1) And i <> d(2) And i <> d(3)) Then
    d(4) = i
    
    Exit For
    
    
    End If
    Next i
    
    
    
    Randomize
    g = Int(Rnd * (7 - 1 + 1)) + 1
    
    Randomize
    i = Int(Rnd * (4 - 1 + 1)) + 1
    Randomize
    u = Int(Rnd * (4 - 1 + 1)) + 1
    Randomize
    Y = Int(Rnd * (4 - 1 + 1)) + 1
    Randomize
    t = Int(Rnd * (4 - 1 + 1)) + 1
    If (g > 2) Then
    Label1.FontSize = 35
    Label1.Caption = Str(i) + " +" + Str(u) + " +" + Str(Y) + " = ?"
    s = i + u + Y
    Else
    Label1.FontSize = 35
    Label1.Caption = Str(i) + " +" + Str(u) + " +" + Str(Y) + " +" + Str(t) + " = ?"
    s = i + u + Y + t
    End If
    
    
    Label2(d(0)).FontSize = 25
    Label2(d(1)).FontSize = 25
    Label2(d(2)).FontSize = 25
    Label2(d(3)).FontSize = 25
    Label2(d(4)).FontSize = 25
    Label2(d(0)).Caption = Str(s)
    Label2(d(1)).Caption = Str(s + 1)
    Label2(d(2)).Caption = Str(s + 2)
    If (g > 3) Then
    
    Label2(d(3)).Caption = Str(s + 3)
    Else
    Label2(d(3)).Caption = Str(s - 2)
    End If
    Label2(d(4)).Caption = Str(s - 1)
    
    
End Sub



Private Sub Label2_Click(Index As Integer)
    
    pi = Int(Label2(Index).Caption)
    If (pi = s) Then
    Unload Form2
    Form2.Show
    
    
    Else
    Label1.BackColor = vbRed
    Label2(Index).BackColor = &H8000000F
    End If
    
End Sub



Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2(Index).BackColor = vbGrayText
End Sub

Private Sub Timer1_Timer()
    Label3.BackColor = vbRed
    Label3.Caption = "time out"
    Label3.ForeColor = vbBlack
End Sub






















