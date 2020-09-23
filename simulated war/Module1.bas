Attribute VB_Name = "Module1"
Public Type unit
name As String
team As Integer
color As Long

max_hp As Long
current_hp As Long
dammage As Long
armor As Long
speed As Long

status As String
current_x As Long
current_y As Long
current_z As Long

target As Integer
range As Integer

attack_rate As Integer
selected As Boolean
End Type

Public Type nearest
id As Integer
distance As Integer
End Type

Public Type sprite
team(1 To 3, 1 To 16, 1 To 16) As Long
End Type
