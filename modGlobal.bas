Attribute VB_Name = "modGlobal"
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

' Memory methods and thread control
Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' GDI methods
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const PS_SOLID = 0

Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Type RECT
    Left As Long
    TOP As Long
    Right As Long
    Bottom As Long
End Type

Type RGBQUAD
    rgbBlue As Byte
    rgbgreen As Byte
    rgbred As Byte
    rgbReserved As Byte
End Type

Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0) As SAFEARRAYBOUND
End Type

  '**********************************************************
   ' Declarations section of the module
   '**********************************************************
   '==========================================================
   ' The following function is designed for use in the AfterUpdate
   ' property of form controls.
   ' Features:
   '    - Leading spaces do not affect the function's performance.
   '    - "O'Brian" and "Wilson-Smythe" will be properly capitalized.
   ' Limitations:
   '    - It will change "MacDonald" to "Macdonald."
   '    - It will change "van Buren" to "Van Buren."
   '    - It will change "John Jones III" to "John Jones Iii."
   '==========================================================
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function OpenFile(FilePath, OwnerHWnd As Long, StartupDirectory As String, nShowCmd As Long) As Long
    OpenFile = ShellExecute(OwnerHWnd, "Open", FilePath, vbNullString, StartupDirectory, nShowCmd)
End Function

   Function Proper(AnyValue As Variant) As Variant
      Dim ptr As Integer
      Dim TheString As String
      Dim currChar As String, prevChar As String

      If IsNull(AnyValue) Then
         Exit Function
      End If

      TheString = CStr(AnyValue)
      For ptr = 1 To Len(TheString)         ' Go through each char. in
                                            ' string.
      currChar = Mid$(TheString, ptr, 1)    ' Get the current character.

         Select Case prevChar               ' If previous char. is a
                                            ' letter,'this char. should be
                                            ' lowercase.
         Case "A" To "Z", "a" To "z"
            Mid(TheString, ptr, 1) = LCase(currChar)

         Case Else
            Mid(TheString, ptr, 1) = UCase(currChar)

      End Select
      prevChar = currChar
      Next ptr
      AnyValue = TheString
   End Function

Public Function SafeUBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long

    On Error Resume Next
    Dim lAddress&, cElements&, lLbound&, cDims%
    
    If Dimension < 1 Then
        SafeUBound = -1
        Exit Function
    End If
    
    CopyMemory lAddress, ByVal lpArray, 4
    
    If lAddress = 0 Then
        ' The array isn't initilized
        SafeUBound = -1
        Exit Function
    End If
    
    ' Calculate the dimenstions
    CopyMemory cDims, ByVal lAddress, 2
    Dimension = cDims - Dimension + 1
    
    ' Obtain the needed data
    CopyMemory cElements, ByVal (lAddress + 16 + ((Dimension - 1) * 8)), 4
    CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4
    
    SafeUBound = cElements + lLbound - 1

End Function

Public Function SafeLBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long

    On Error Resume Next
    Dim lAddress&, cElements&, lLbound&, cDims%
    
    If Dimension < 1 Then
        SafeLBound = -1
        Exit Function
    End If
    
    CopyMemory lAddress, ByVal lpArray, 4
    
    If lAddress = 0 Then
        ' The array isn't initilized
        SafeLBound = -1
        Exit Function
    End If
    
    ' Calculate the dimenstions
    CopyMemory cDims, ByVal lAddress, 2
    Dimension = cDims - Dimension + 1
    
    ' Obtain the needed data
    CopyMemory lLbound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4
    
    SafeLBound = lLbound

End Function

