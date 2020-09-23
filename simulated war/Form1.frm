VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   Caption         =   "AI WAR SIM"
   ClientHeight    =   12090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14745
   DrawStyle       =   2  'Dot
   DrawWidth       =   5
   LinkTopic       =   "Form1"
   ScaleHeight     =   806
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   983
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Text            =   "1.0025"
         ToolTipText     =   "EVOLUTION MULTIPLIER"
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Blood cell"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pause"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000001&
         ForeColor       =   &H0000FF00&
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Text            =   "10"
         ToolTipText     =   "# of AI UNITS"
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Low Graphics"
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Lkills 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Lkills 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Lkills 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000008&
         Caption         =   "Label1"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a() As unit, current As Integer, i As Integer, b As nearest, units As Integer, ok, aaa(4) As Long, aax As Long, aay As Long, k As sprite, li As Integer, tkills(1 To 3)

Private Sub Command1_Click()
tkills(1) = 0: tkills(2) = 0: tkills(3) = 0
Text1.Text = ""
Timer1.Enabled = False
hhh
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Pause" Then: Timer1.Enabled = False: Command2.Caption = "Resume": Else Timer1.Enabled = True: Command2.Caption = "Pause"
End Sub

Private Sub Form_Load()
get_sprites
hhh
End Sub
Private Sub hhh()
units = Val(Text2.Text)
ReDim a(units)
generate_units
End Sub
Private Sub generate_units()
Randomize Timer
For i = 0 To units - 1
a(i).status = "Born"
a(i).name = "player " & i
a(i).armor = 2 + Int(Rnd * 2) * 3
a(i).max_hp = 450 + (Int(Rnd * 5) * 200)
a(i).current_hp = a(i).max_hp
a(i).dammage = 4 + (Int(Rnd * 2) * 8)

'If i >= units / 2 Then a(i).team = 2 Else a(i).team = 1
a(i).team = Int(Rnd * 3) + 1
If a(i).team = 1 Then a(i).color = vbYellow
If a(i).team = 2 Then a(i).color = vbGreen
If a(i).team = 3 Then a(i).color = vbRed
a(i).current_x = Int(Rnd * Screen.Width / Screen.TwipsPerPixelX)
a(i).current_y = Int(Rnd * Screen.Height / Screen.TwipsPerPixelY)
a(i).speed = 1 + ((2 + (Rnd * 10)) / 2)
a(i).range = 20 + Int(Rnd * 15)
Next i
End Sub
Private Sub render()
If Check2.Value = 1 Then Cls
For i = 0 To units - 1
If a(i).status <> "dead" Then draw
Next i

'For i = 0 To 7
'Label1(i) = a(i).current_hp
'Next
For op = 1 To 3
Lkills(op - 1) = tkills(op)
Next op
End Sub
Private Sub draw()
aax = a(i).current_x: aay = a(i).current_y
If Check1.Value <> 1 Then draw_sprite Else: Me.DrawWidth = 1: Circle (a(i).current_x, a(i).current_y), 7, a(i).color
'If a(i).status = "attacking" Then Circle (a(i).current_x, a(i).current_y), 7, vbRed Else _
'Circle (a(i).current_x, a(i).current_y), 3, vbGreen
'Caption = a(5).status & " " & a(5).current_hp
'Circle (a(5).current_x, a(5).current_y), 10, vbGreen
End Sub
Private Sub bmp() ' not used (old bmp loader) inputs file for every object = bad!
Seek a(i).team, 55
For aaj = 1 To 16
For aai = 1 To 16
stuf = Input(3, a(i).team)
If stuf = "" Then Exit Sub
aaa(0) = Asc(Mid(stuf, 1, 1))
aaa(1) = Asc(Mid(stuf, 2, 1))
aaa(2) = Asc(Mid(stuf, 3, 1))
acolor = RGB(aaa(2), aaa(1), aaa(0))
If acolor <> vbBlack Then Circle (aai + (aax - awidth / 2), (aaj * -1) + aay), 10, acolor
Next aai
Next aaj
End Sub
Private Sub get_sprites()
Open "test.bmp" For Binary As #1
Open "test1.bmp" For Binary As #2
Open "test2.bmp" For Binary As #3
For aateam = 1 To 3 'team #
Seek aateam, 55     'boRGBstream
For aaj = 1 To 16
For aai = 1 To 16
stuf = Input(3, aateam) 'inputs 3 bytes RGB
If stuf = "" Then Exit Sub
aaa(0) = Asc(Mid(stuf, 1, 1))
aaa(1) = Asc(Mid(stuf, 2, 1))
aaa(2) = Asc(Mid(stuf, 3, 1))
acolor = RGB(aaa(2), aaa(1), aaa(0))
If acolor <> vbBlack Then k.team(aateam, aai, aaj) = acolor 'stores picture in array k(team,x,y) ,excluding Black (clear color)
Next aai
Next aaj
Next aateam
Close #1
Close #2
Close #3
End Sub
Private Sub draw_sprite() 'retrives picture from array
If Check3.Value = 1 Then Me.DrawWidth = 2: s = 10 Else Me.DrawWidth = 3: s = 1
For aaj = 1 To 16 'Y
For aai = 1 To 16 'X
If k.team(a(i).team, aai, aaj) <> 0 Then Circle (-10 + aai + (aax - awidth / 2), 10 + (aaj * -1) + aay), s, k.team(a(i).team, aai, aaj)
Next aai
Next aaj
End Sub
Private Sub prosses()
For i = 0 To units - 1
current = i
Select Case a(i).status
Case "attacking"
attack
Case "walking"
walk
Case "Born"
bored
Case "dead"
Case ""
bored
Case Else
bored
End Select
Next i
is_all_dead
render

End Sub
Private Sub attack()
'Me.FillColor = vbRed
If a(current).target = -1 Then a(current).status = "Born": Exit Sub
If a(current).status = "dead" Then Exit Sub
If a(a(current).target).status = "dead" Then: a(current).status = "Born": Exit Sub
distance = Sqr((a(current).current_x - a(a(current).target).current_x) ^ 2 + (a(current).current_y - a(a(current).target).current_y) ^ 2)
If distance > a(current).range Then a(current).status = "born": Exit Sub
a(a(current).target).target = current

If Check1.Value = 1 Then Me.DrawWidth = 1 Else Me.DrawWidth = 5
Line (a(a(current).target).current_x, a(a(current).target).current_y)-(a(current).current_x, a(current).current_y), vbRed 'a(current).color 'vbRed
'Circle (a(current).current_x, a(current).current_y), 20, vbRed
DoEvents

a(a(current).target).current_hp = a(a(current).target).current_hp - (a(current).dammage + a(a(current).target).armor)


If a(a(current).target).current_hp < 0 Then a(a(current).target).status = "dead": a(current).status = "Born": _
Text1.Text = Text1.Text & vbNewLine & a(current).name & " killed " & a(a(current).target).name: death_flash: respawn_dead: _
give_power

Label1(7) = a(current).current_hp

End Sub
Private Sub death_flash()
If Check1.Value = 1 Then Exit Sub
For i = 1 To 20
ooo = 255 - ((i / 20) * 254)
Me.DrawWidth = 2
Circle (a(a(current).target).current_x, a(a(current).target).current_y), i, RGB(ooo, 0, 0)
DoEvents
Next i
End Sub
Private Sub give_power()
tkills(a(current).team) = tkills(a(current).team) + 1

mult = Val(Text3.Text)
Select Case Int(Rnd * 6) + 1
Case 1
a(current).speed = a(current).speed * mult
If a(current).speed > 10 Then a(current).speed = 10
Case 2
a(current).dammage = a(current).dammage * mult
If a(current).dammage > (10 ^ 9) Then a(current).dammage = (10 ^ 9)
Case 3
a(current).max_hp = a(current).max_hp * mult
If a(current).max_hp > (10 ^ 9) Then a(current).max_hp = (10 ^ 9)
Case 4
a(current).armor = a(current).armor * mult
If a(current).armor > (10 ^ 9) Then a(current).armor = (10 ^ 9)
Case 5
a(current).range = a(current).range * mult
If a(current).range > 75 Then a(current).range = 75
Case Else
a(current).dammage = a(current).dammage * mult
If a(current).dammage > (10 ^ 9) Then a(current).dammage = (10 ^ 9)
End Select
a(current).current_hp = a(current).current_hp + 200

 ': a(current).target = -1
End Sub
Private Sub respawn_dead()
a(current).status = "born": 'a(current).target = -1
li = a(current).target: spawn
'For li = 0 To units - 1
'If a(li).status = "dead" And a(li).team <> a(current).team Then spawn: Exit Sub
'Next li
End Sub
Private Sub spawn()
a(li).status = "Born"
a(li).current_hp = a(li).max_hp
a(li).current_x = Int(Rnd * Screen.Width / Screen.TwipsPerPixelX)
a(li).current_y = Int(Rnd * Screen.Height / Screen.TwipsPerPixelY)

If Check1.Value = 1 Then Exit Sub
For i = 1 To 20
ooo = 255 - ((i / 20) * 254)
Me.DrawWidth = 2
Circle (a(li).current_x, a(li).current_y), i * 2, RGB(0, ooo, 0)
DoEvents
Next i
End Sub
Private Sub is_all_dead()

End Sub
Private Sub walk()
If a(current).target = -1 Then find_nearest: a(current).status = "Born": Exit Sub
If a(a(current).target).status = "dead" Then find_nearest: a(current).status = "walking": a(current).target = b.id: Exit Sub
If a(current).target = -1 Then MsgBox "the war is over": End
'select case
If a(current).current_x < a(a(current).target).current_x Then a(current).current_x = a(current).current_x + a(current).speed
If a(current).current_x > a(a(current).target).current_x Then a(current).current_x = a(current).current_x - a(current).speed
If a(current).current_y < a(a(current).target).current_y Then a(current).current_y = a(current).current_y + a(current).speed
If a(current).current_y > a(a(current).target).current_y Then a(current).current_y = a(current).current_y - a(current).speed

distance = Sqr((a(current).current_x - a(a(current).target).current_x) ^ 2 + (a(current).current_y - a(a(current).target).current_y) ^ 2)
If distance <= a(current).range Then a(current).status = "attacking"
End Sub
Private Sub bored()
find_nearest
a(current).status = "walking"
a(current).target = b.id
End Sub
Private Sub find_nearest()
b.distance = 32000
b.id = -1
For i = 0 To units - 1
ok = i
If i <> current And a(current).team <> a(i).team And a(i).status <> "dead" Then compare
Next i
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Frame1.Visible = True Then Frame1.Visible = False Else Frame1.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
Close #2
Close #3
End
End Sub
Private Sub Timer1_Timer()
prosses
End Sub
Private Sub compare()
distance = Sqr((a(ok).current_x - a(current).current_x) ^ 2 + (a(ok).current_y - a(current).current_y) ^ 2)
If distance < b.distance Then b.distance = distance: b.id = ok
End Sub
