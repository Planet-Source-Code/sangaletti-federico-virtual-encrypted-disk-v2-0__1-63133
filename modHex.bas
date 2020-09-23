Attribute VB_Name = "modHEX"
'-------------------------------------------------------
'   Project..... VIRTUAL ENCRYPTED DISK UTILITY v2.0
'   Module...... modHEX
'   Author...... Fredrik Qvarfort
'   License..... FREE (but respect copyright)
'
'   Decription.. This module is part of clsCryptAPI
'                and contains functions to encode an
'                hex number to corrispondent string
'                and vice-versa
'-------------------------------------------------------

Option Explicit

Private m_InitHex As Boolean
Private m_ByteToHex(0 To 255, 0 To 1) As Byte
Private m_HexToByte(48 To 70, 48 To 70) As Byte

Function HexToStr(HexText As String) As String

  Dim a As Long
  Dim POS As Long
  Dim ByteSize As Long
  Dim HexByte() As Byte
  Dim ByteArray() As Byte
  
  'Initialize the hex routine
  If (Not m_InitHex) Then Call InitHex
  
  'The destination string is half
  'the size of the source string
  'when the separators are removed
  If (Len(HexText) = 2) Then
    ByteSize = 1
  Else
    ByteSize = (Len(HexText) + 1) \ 2
  End If
  ReDim ByteArray(0 To ByteSize - 1)
  
  'Convert every HEX code to the
  'equivalent ASCII character
  HexByte() = StrConv(HexText, vbFromUnicode)
  For a = 0 To (ByteSize - 1)
    ByteArray(a) = m_HexToByte(HexByte(POS), HexByte(POS + 1))
    POS = POS + 2
  Next
  
  'Now finally convert the byte
  'array to the return string
  HexToStr = StrConv(ByteArray, vbUnicode)

End Function
Private Sub InitHex()

  Dim a As Long
  Dim B As Long
  Dim HexBytes() As Byte
  Dim HexString As String
  
  'The routine is initialized
  m_InitHex = True
  
  'Create a string with all hex values
  HexString = String$(512, "0")
  For a = 1 To 255
    Mid$(HexString, 1 + a * 2 + -(a < 16)) = Hex(a)
  Next
  HexBytes = StrConv(HexString, vbFromUnicode)
  
  'Create the Str->Hex array
  For a = 0 To 255
    m_ByteToHex(a, 0) = HexBytes(a * 2)
    m_ByteToHex(a, 1) = HexBytes(a * 2 + 1)
  Next
  
  'Create the Str->Hex array
  For a = 0 To 255
    m_HexToByte(m_ByteToHex(a, 0), m_ByteToHex(a, 1)) = a
  Next

End Sub
Function StrToHex(Text As String) As String

  Dim a As Long
  Dim POS As Long
  Dim ByteSize As Long
  Dim ByteArray() As Byte
  Dim ByteReturn() As Byte
  
  'Initialize the hex routine
  If (Not m_InitHex) Then Call InitHex
  
  'Create the destination bytearray, this
  'will be converted to a string later
  ByteSize = Len(Text) * 2 + (Len(Text) - 1)
  ReDim ByteReturn(ByteSize - 1)
  
  'We convert the source string into a
  'byte array to speed this up a tad
  ByteArray() = StrConv(Text, vbFromUnicode)
  
  'Now convert every character to
  'it's equivalent HEX code
  For a = 0 To (Len(Text) - 1)
    ByteReturn(POS) = m_ByteToHex(ByteArray(a), 0)
    ByteReturn(POS + 1) = m_ByteToHex(ByteArray(a), 1)
    POS = POS + 2
  Next
  
  'Convert the bytearray to a string
  StrToHex = StrConv(ByteReturn(), vbUnicode)

End Function
