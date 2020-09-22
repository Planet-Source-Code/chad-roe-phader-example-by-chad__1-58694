Attribute VB_Name = "PhadeBas"
Option Explicit

Public Sub Hex2RGB(strHexColor As String, R As Byte, G As Byte, B As Byte)
On Error Resume Next
Dim i As Byte, HexColor As String
strHexColor = Right((strHexColor), 6)
 For i = 1 To (6 - Len(strHexColor))
  HexColor = HexColor & "0"
 Next
 HexColor = HexColor & strHexColor
 R = CByte("&H" & Right$(HexColor, 2))
 G = CByte("&H" & Mid$(HexColor, 3, 2))
 B = CByte("&H" & Left$(HexColor, 2))
End Sub

Public Function RGB2Hex(R As Byte, G As Byte, B As Byte) As String
On Error Resume Next
RGB2Hex = Long2Hex(RGB(R, G, B))
End Function

Public Sub Long2RGB(LongColor As Long, R As Byte, G As Byte, B As Byte)
On Error Resume Next
Hex2RGB (Hex(LongColor))
End Sub

Public Function RGB2Long(R As Byte, G As Byte, B As Byte) As Long
On Error Resume Next
RGB2Long = RGB(R, G, B)
End Function

Public Function Long2Hex(LongColor As Long) As String
On Error Resume Next
Long2Hex = Hex(LongColor)
End Function

Public Function Hex2Long(strHexColor As String) As Long
Dim R As Byte, G As Byte, B As Byte
On Error Resume Next
Call Hex2RGB(strHexColor, R, G, B)
Hex2Long = RGB(R, G, B)
End Function
