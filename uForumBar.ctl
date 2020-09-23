VERSION 5.00
Begin VB.UserControl uSideBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   184
   ToolboxBitmap   =   "uForumBar.ctx":0000
   Begin VB.VScrollBar vscOverflow 
      Height          =   2295
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "uSideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Enum BarState
    BarState_Collapsed
    BarState_Expanded
End Enum

Public Enum EventCodes
    BarEvent_Changing
    BarEvent_Changed
    BarEvent_ItemAdded
    BarEvent_ItemRemoved
    BarEvent_ItemChanged
End Enum

' Different settings regarding bars
Public Enum BarTitle
    BarTitle_Width = 185
    bartitle_height = 23
    BarTitle_CornerLenght = 3
    BarTitle_Left = 13
    BarTitle_BeginTop = 13
    BarTitle_Space = 15
    BarTitle_UpAndDownSpace = 10
    BarTitle_ItemLeft = 10
    BarTitle_ItemSpace = 2
    BarTitle_IconSpace = 5
End Enum

' GDI-const
Enum DT_CONST
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
End Enum

Private Type RGB
    Red As Byte
    Green As Byte
    Blue As Byte
    Unused As Byte
End Type

Event Click(ClickedObject As Object)
Event Redrawed()
Event Redrawing(Cancel As Boolean)
Event Changing()
Event Changed(ChangeString As String, ChangedObject As Object, ExtraInfo As String)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

' Public variables
Public IgnoreRedraw As Boolean

Dim SelectedItem As Long
Dim Bars() As clsBarNode

Public Property Get Bar(Index As Variant) As clsBarNode
    
    Dim Tell As Long
    
    If IsNumeric(Index) Then
        If Index < BarStartIndex Or Index > BarCount Then Exit Property
        Set Bar = Bars(Index)
    Else
    
        For Tell = 0 To BarCount
        
            If LCase(Index) = LCase(Bars(Tell).ItemData) Then
                Set Bar = Bars(Tell)
                Exit Property
            End If
        
        Next
    
    End If
    
End Property

Public Property Get BarCount() As Long

    BarCount = SafeUBound(VarPtrArray(Bars))
    
End Property

Public Property Get BarSelected() As Long

    BarSelected = SelectedItem
    
End Property

Public Property Get BarStartIndex() As Long

    BarStartIndex = 0 ' Always 0
    
End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = UserControl.BackColor
    
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)

    UserControl.BackColor = vNewValue
    
    ' After changing the backcolor, redraw all bars and controls
    Redraw
    
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", RGB(123, 162, 231))
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "BackColor", UserControl.BackColor, RGB(123, 162, 231)
    
End Sub

Public Sub HandleEvent(eCode As EventCodes, ParamArray eParamenters() As Variant)

    Dim lpObject As Object, lpRefObject As Object, sTemp As String
    
    Select Case eCode
    Case BarEvent_Changing
    
        ' Tell about this changing
        RaiseEvent Changing
    
    Case BarEvent_Changed, BarEvent_ItemAdded, BarEvent_ItemRemoved, BarEvent_ItemChanged
    
        If eCode = BarEvent_ItemAdded Then
            If Bars(eParamenters(0)).Item(eParamenters(1)).ClassMember = ItemType_Object Then
                ' No point of redrawing, since it's an remote object
                Exit Sub
            End If
        End If
    
        If eCode = BarEvent_ItemChanged Then
        
            If eParamenters(1) = "ItemData" Or eParamenters(1) = "Tag" Then
                ' We don't need to redraw anything when these event occures
                Exit Sub
            End If
            
            If eParamenters(1) = "Control" Then
    
                Set lpObject = eParamenters(0).Control
                
                ' See if we need to redraw
                If Not lpObject Is Nothing Then
                    Redraw
                End If
                
            Else
        
                Set lpObject = eParamenters(0)
                DrawItem lpObject, lpObject.Index, lpObject.Parent.ItemPositionY(lpObject.Index), True
            
            End If
            
        Else
            ' Redraw all the controls
            Redraw
        End If
        
        ' Get the object that is the subject for changing
        If eCode = BarEvent_Changed Then
            Set lpRefObject = Bars(Val(eParamenters(0)))
            
        ElseIf eCode = BarEvent_ItemChanged Then
            
            ' The first paramenter is the object being changed
            Set lpRefObject = eParamenters(0)
            
            ' The next paramenter is what's being changed in the object
            sTemp = eParamenters(1)
            
        Else
            Set lpRefObject = Bars(Val(eParamenters(0))).Item(Val(eParamenters(1)))
            
        End If
        
        ' Raise the event
        RaiseEvent Changed(Choose(eCode, "BarChanged", "ItemAdded", "ItemRemoved", "ItemChanged"), lpRefObject, sTemp)
        
    End Select

End Sub

Public Sub RemoveBar(ByVal Index As Long)

    Dim Tell As Long
    
    If Index < BarStartIndex Or Index > BarCount Then
        Exit Sub
    End If
    
    Set Bars(Index) = Nothing
    
    For Tell = Index To BarCount - 1
        CopyMemory ByVal VarPtr(Bars(Tell)), ByVal VarPtr(Bars(Tell + 1)), 4
    Next
    
    ' Reallocate the array to its right size.
    If UBound(Bars) > 0 Then
        ReDim Preserve Bars(UBound(Bars) - 1)
    Else
        Erase Bars
    End If
    
    ' Redraw all the controls
    Redraw

End Sub

' Add bars
Public Function AddBar(Optional Key As String, Optional Title As String) As clsBarNode

    ' Realllocate the array
    ReDim Preserve Bars(SafeUBound(VarPtrArray(Bars)) + 1)
    
    ' Create a new bar
    Set Bars(UBound(Bars)) = New clsBarNode
    Set Bars(UBound(Bars)).Parent = Me
    
    Bars(UBound(Bars)).Caption = Title
    Bars(UBound(Bars)).ItemData = Key
    Bars(UBound(Bars)).Tag = Title
    Bars(UBound(Bars)).Index = UBound(Bars)
    Bars(UBound(Bars)).TitleFont = New StdFont
    Bars(UBound(Bars)).TitleFont.Bold = True
    Bars(UBound(Bars)).State = BarState_Expanded
    Bars(UBound(Bars)).InvokeEvents = True
    
    ' Return the created object
    Set AddBar = Bars(UBound(Bars))
    
    ' Redraw all the controls
    Redraw

End Function

Public Property Get BarExists(ByVal Index As Variant) As Boolean

    Dim Tell As Long
    
    If IsNumeric(Index) Then
    
        If Index < BarStartIndex Or Index > BarCount Then
            BarExists = False
            Exit Property
        End If
        
        If Bars(Index) Is Nothing Then
            BarExists = False
            Exit Property
        End If
    
    Else
    
        For Tell = 0 To BarCount
        
            If LCase(Index) = LCase(Bars(Tell).ItemData) Then
                BarExists = True
                Exit Property
            End If
        
        Next
    
        Exit Property
    End If
    
    BarExists = True

End Property

Public Sub DeselectAllItems()

    Dim Tell As Long
    
    For Tell = BarStartIndex To BarCount
        Bars(Tell).DeselectItems
    Next

End Sub

Public Sub SelectSingleItem(BarIndex As Long, ItemIndex As Long)

    If Bars(BarIndex).Item(ItemIndex).Selected = False Then
        ' Deselect ALL items
        DeselectAllItems
        
        ' Select that item
        Bars(BarIndex).Item(ItemIndex).Selected = True
    End If

End Sub

' This is the draw-routine
Public Sub Redraw()

    Dim Tell As Long, Tmp As Long, CurrY As Long, ItemY As Long, bCancel As Boolean
    Dim lMax As Long, lValue As Long, tempIgnore As Boolean, ScrollIndex As Long
    
    If IgnoreRedraw Then Exit Sub
    RaiseEvent Redrawing(bCancel)
    If bCancel = True Then Exit Sub
    
    CurrY = BarTitle_BeginTop
    
    ' Paint over the last image-state
    UserControl.Cls
    
    For Tell = BarStartIndex To BarCount
        ' Give the buffer the same font as the bar
        UserControl.ForeColor = IIf(Tell = SelectedItem, Bars(Tell).TitleForeColorOver, Bars(Tell).TitleForeColor)
        UserControl.FontBold = Bars(Tell).TitleFont.Bold
        UserControl.FontItalic = Bars(Tell).TitleFont.Italic
        UserControl.FontName = Bars(Tell).TitleFont.Name
        UserControl.FontSize = Bars(Tell).TitleFont.Size
        UserControl.FontUnderline = Bars(Tell).TitleFont.Underline
        
        DrawCorneredRect BarTitle_Left, CurrY, BarTitle_Width, bartitle_height, BarTitle_CornerLenght, Bars(Tell).TitleBackColorLight, Bars(Tell).TitleBackColorDark
        DrawText UserControl.hdc, Bars(Tell).Caption, BarTitle_Left + BarTitle_CornerLenght + 5, CurrY, BarTitle_Width - BarTitle_CornerLenght - 3, bartitle_height, DT_VCENTER Or DT_NOCLIP Or DT_SINGLELINE
        
        ' Draw the button
        DrawCircle UserControl.hdc, BarTitle_Width + BarTitle_Left - 19 - 5, CurrY + 2, 20, 20, Bars(Tell).ButtonBorder, Bars(Tell).ButtonFace, Bars(Tell).ButtonShadow
        DrawText UserControl.hdc, IIf(Bars(Tell).State = BarState_Collapsed, "«", "»"), BarTitle_Width + BarTitle_Left - 19 - 5, CurrY, 19, 19, DT_VCENTER Or DT_NOCLIP Or DT_SINGLELINE + DT_CENTER
        
        ' Clean up the height from last time
        Bars(Tell).Height = BarTitle_UpAndDownSpace * 2
        
        ' ItemY, which indicates the current Y position where each new object is drawed,
        ' must be set to default each time we're about to draw all items.
        ItemY = BarTitle_UpAndDownSpace + CurrY + bartitle_height
        
        ' We also need to find out whether or not the scroll-bar is necessary
        If Bars(Tell).ItemCount > Bars(Tell).MaxItems And Bars(Tell).MaxItems > 0 And Bars(Tell).State = BarState_Expanded Then
        
            If Bars(Tell).ScrollControl Is Nothing Then
                ' Find out what index the new scroll should have
                ScrollIndex = GetScrollIndex
                
                ' See if scrollvalue is valid
                If Bars(Tell).ScrollValue > Bars(Tell).ItemCount Or Bars(Tell).ScrollValue < 0 Then
                    Bars(Tell).ScrollValue = 0
                End If
                
                ' Save the state
                tempIgnore = IgnoreRedraw
                
                IgnoreRedraw = True
                vscOverflow(ScrollIndex).Left = BarTitle_Left + BarTitle_Width - BarTitle_ItemLeft - vscOverflow(vscOverflow.Count - 1).Width
                vscOverflow(ScrollIndex).TOP = ItemY
                vscOverflow(ScrollIndex).Max = Bars(Tell).ItemCount - Bars(Tell).MaxItems
                vscOverflow(ScrollIndex).Value = Bars(Tell).ScrollValue
                vscOverflow(ScrollIndex).Visible = True
        
                ' Restore state
                IgnoreRedraw = tempIgnore
                
                Set Bars(Tell).ScrollControl = vscOverflow(ScrollIndex)
                
            Else
                Bars(Tell).ScrollControl.TOP = ItemY
            End If
            
        Else
            ' If not, then reset ScrollControl and hide the control
            If Not Bars(Tell).ScrollControl Is Nothing Then
                Bars(Tell).ScrollControl.Visible = False
                Set Bars(Tell).ScrollControl = Nothing
            End If
            
        End If
        
        ' Find out of the height
        For Tmp = Bars(Tell).ItemStartIndex To Bars(Tell).ItemCount
            
            If Not Bars(Tell).ScrollControl Is Nothing Then
                If Tmp > Bars(Tell).MaxItems Then
                    ' The height must NOT be bigger than it should be
                    Bars(Tell).ScrollControl.Height = Bars(Tell).Height - (BarTitle_UpAndDownSpace * 2)
                    Exit For
                End If
            End If
            
            Bars(Tell).Height = Bars(Tell).Height + Bars(Tell).Item(Tmp).Height + BarTitle_ItemSpace
        Next
        
        ' Draw the bar
        If Bars(Tell).State = BarState_Expanded Then
            UserControl.Line (BarTitle_Left, CurrY + bartitle_height)-(BarTitle_Left + BarTitle_Width, CurrY + bartitle_height + Bars(Tell).Height), Bars(Tell).BackColor, BF
            UserControl.Line (BarTitle_Left, CurrY + bartitle_height)-(BarTitle_Left + BarTitle_Width, CurrY + bartitle_height + Bars(Tell).Height), Bars(Tell).LineColor, B
            
            ' We are going to draw different things at different states
            If Bars(Tell).ScrollControl Is Nothing Then
                lMax = Bars(Tell).ItemCount
                lValue = 0
            Else
                lMax = Bars(Tell).MaxItems
                lValue = Bars(Tell).ScrollControl.Value
            End If
        
            ' Then draw all the controls
            For Tmp = Bars(Tell).ItemStartIndex + lValue To Bars(Tell).ItemStartIndex + lValue + lMax
                DrawItem Bars(Tell).Item(Tmp), Tmp, ItemY
            Next
        
            CurrY = CurrY + Bars(Tell).Height
        Else
            
            ' If not, then we need to hide all controls associated with this bar.
        
            For Tmp = Bars(Tell).ItemStartIndex To Bars(Tell).ItemCount
                If Bars(Tell).Item(Tmp).ClassMember = ItemType_Object Then
                    If Not Bars(Tell).Item(Tmp).Control Is Nothing Then
                        Bars(Tell).Item(Tmp).Control.Visible = False
                    End If
                End If
            Next
        
        End If
    
        CurrY = CurrY + bartitle_height + BarTitle_Space
    Next
    
    RaiseEvent Redrawed

End Sub

Private Sub DrawItem(Item As Object, Index As Long, ItemY As Long, Optional EmptyArea As Boolean)

    Dim ItemX As Long
    
    ' Holds the X-position of the object
    ItemX = BarTitle_ItemLeft + BarTitle_Left
    
    ' Set to current font
    UserControl.FontBold = Item.Bold
    
    If Item.Selected = True And Item.ClassMember = ItemType_Link Then
        UserControl.ForeColor = Item.TextColorOver
        UserControl.FontUnderline = True
    Else
        UserControl.ForeColor = Item.TextColor
        UserControl.FontUnderline = False
    End If
    
    ' If we must refill the area, then do it
    If EmptyArea = True Then
        UserControl.Line (ItemX, ItemY)-(ItemX + BarTitle_Width - BarTitle_ItemLeft - Item.SpacingAfter - 1, ItemY + Item.Height), Item.Parent.BackColor, BF
    End If
    
    ' Draw the icon
    If Not Item.IconHandle Is Nothing Then
        UserControl.PaintPicture Item.IconHandle, ItemX, ItemY
        ItemX = ItemX + UserControl.ScaleX(Item.IconHandle.Width, vbHimetric, vbPixels) + BarTitle_IconSpace
    End If
    
    Select Case Item.ClassMember
    Case ItemType_Link, ItemType_ControlPlaceHolder
    
        ' Center the text vertically
        DrawText UserControl.hdc, Item.Text, ItemX, ItemY, BarTitle_Width - BarTitle_ItemLeft - Item.SpacingAfter, Item.Height, DT_VCENTER + DT_WORDBREAK
    
    Case ItemType_Object
    
        Item.Control.Left = ItemX
        Item.Control.TOP = ItemY + 50
        Item.Control.Visible = True
        
    End Select
    
    ItemY = ItemY + Item.Height + BarTitle_ItemSpace

End Sub

Private Function DrawText(ByVal hdc As Long, sText As String, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal wFormat As DT_CONST) As Long

    Dim rRECT As RECT
    
    With rRECT
        .Left = lX
        .TOP = lY
        .Right = lX + lWidth
        .Bottom = lY + lHeight
    End With
    
    DrawText = DrawTextA(hdc, sText, Len(sText), rRECT, wFormat)

End Function

Private Function DrawCorneredRect(ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lCornerLenght As Long, ByVal lBeginColor As Long, ByVal lEndColor As Long) As Long

    Dim tempX As Long, tempY As Long, lColorStep As Long, point&, tempColor As Long
    Dim hPen As Long, beginRGB As RGB, endRGB As RGB, stepRed#, stepGreen#, stepBlue#
    
    CopyMemory beginRGB, lBeginColor, 4
    CopyMemory endRGB, lEndColor, 4
    
    stepRed = (CLng(endRGB.Red) - CLng(beginRGB.Red)) / lWidth
    stepGreen = (CLng(endRGB.Green) - CLng(beginRGB.Green)) / lWidth
    stepBlue = (CLng(endRGB.Blue) - CLng(beginRGB.Blue)) / lWidth
    
    For tempX = lX To lX + lWidth
    
        tempY = lY
        
        point = tempX - lX
        tempColor = RGB(CLng(beginRGB.Red) + (stepRed * point), CLng(beginRGB.Green) + (stepGreen * point), CLng(beginRGB.Blue) + (stepBlue * point))
    
        If tempX - lX < lCornerLenght Then tempY = lY + (lCornerLenght - (tempX - lX))
        If tempX - lX - lWidth > -lCornerLenght Then tempY = lY + (lCornerLenght + (tempX - lX - lWidth))
        
        UserControl.Line (tempX, tempY)-(tempX, lY + lHeight), tempColor
        
    Next

End Function

Private Function DrawCircle(hDestDC As Long, X As Long, Y As Long, Width As Long, Height As Long, BorderColor As Long, ButtonFace As Long, ButtonShadow As Long)

    Dim bi24BitInfo As BITMAPINFO, bBytes() As Byte, bTemp As Long, Tell As Long, iDC As Long, iBitmap As Long
    Dim sa As SAFEARRAY1D, lngPointer As Long, ByteWidth As Long, hBrush As Long, hPen As Long
    
    ' Initialize DIB
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = Width
        .biHeight = Height
    End With
    
    iDC = CreateCompatibleDC(0)
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, lngPointer, ByVal 0&, ByVal 0&)
    SelectObject iDC, iBitmap
    
    ' Fill the surface with the current background
    BitBlt iDC, 0, 0, Width, Height, hDestDC, X, Y, vbSrcCopy
    
    ' Create a brush and a pen for the shadow
    hBrush = CreateSolidBrush(ButtonShadow)
    hPen = CreatePen(PS_SOLID, 1, ButtonShadow)
    SelectObject iDC, hBrush
    SelectObject iDC, hPen
    
    ' Draw the shadow
    Ellipse iDC, 2, 2, Width - 3, Height - 3
    
    ' Free resources
    DeleteObject hPen
    DeleteObject hBrush
        
    ' Create a pen and a brush
    hPen = CreatePen(PS_SOLID, 1, BorderColor)
    hBrush = CreateSolidBrush(ButtonFace)
    SelectObject iDC, hBrush
    SelectObject iDC, hPen
    
    ' Draw the ellipse
    Ellipse iDC, 1, 1, Width - 4, Height - 4
    
    With sa
        .cDims = 1
        .Bounds(0).cElements = Width * Height * 3
        .pvData = lngPointer
    End With
    
    ' Speed up calculations by precaluclate the width of a line in bytes
    ByteWidth = Width * 3
    
    CopyMemory ByVal VarPtrArray(bBytes), VarPtr(sa), 4
    
    For Tell = LBound(bBytes) + ByteWidth + 3 To UBound(bBytes) - ByteWidth - 3
        bBytes(Tell) = (CInt(bBytes(Tell)) + CInt(bBytes(Tell + 3)) + CInt(bBytes(Tell - 3)) + CInt(bBytes(Tell + ByteWidth + 3)) + CInt(bBytes(Tell - ByteWidth - 3))) \ 5
    Next
    
    ' Draw the result to the destination dc
    BitBlt hDestDC, X, Y, Width, Height, iDC, 0, 0, vbSrcCopy
    
    CopyMemory ByVal VarPtrArray(bBytes), 0&, 4
    
    ' Clean up resources
    DeleteDC iDC
    DeleteObject iBitmap
    DeleteObject hBrush
    DeleteObject hPen

End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lastSI As Long
    
    ' Tell about the movement
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    lastSI = SelectedItem
    SelectedItem = GetSelectedItem(X, Y)
    FindSelectedObject X, Y ' Will redraw the item anyway, so we don't need to do that
    
    If lastSI <> SelectedItem Then
        ' Redraw all the controls
        Redraw
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lpItem As Object
    
    ' This event is primary for user-defined action in the control area
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    SelectedItem = GetSelectedItem(X, Y)
    Set lpItem = FindSelectedObject(X, Y)
    
    If SelectedItem >= 0 Then
        Bars(SelectedItem).State = IIf(Bars(SelectedItem).State = BarState_Expanded, BarState_Collapsed, BarState_Expanded)
    End If
    
    ' Here we tell about what has been selected
    If SelectedItem >= 0 Then
        RaiseEvent Click(Bars(SelectedItem))
    ElseIf Not lpItem Is Nothing Then
        RaiseEvent Click(lpItem)
    End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Function FindSelectedObject(X As Single, Y As Single) As Object

    Dim Tell As Long, tempY As Long, lItemCount As Long, BarTell As Long, lValue As Long, lMax As Long
    
    If X > BarTitle_Left + BarTitle_ItemLeft And X < BarTitle_Left + BarTitle_Width Then
    
        For BarTell = BarStartIndex To BarCount
        
            ' Get the total count of items
            lItemCount = Bars(BarTell).ItemCount
    
            ' Find the Y-position of the current selected bar
            tempY = BarPositionY(BarTell) + bartitle_height + (BarTitle_UpAndDownSpace * 2)
    
            If Y < tempY + Bars(BarTell).Height And Y > tempY And Bars(BarTell).State = BarState_Expanded Then
    
                If lItemCount - Bars(BarTell).ItemStartIndex < 0 Then
                    DeselectAllItems
                    Exit Function
                End If
        
                If Bars(BarTell).ScrollControl Is Nothing Then
                    lMax = lItemCount
                    lValue = 0
                Else
                    lMax = Bars(BarTell).MaxItems
                    lValue = Bars(BarTell).ScrollControl.Value
                End If
                
                For Tell = Bars(BarTell).ItemStartIndex + lValue To Bars(BarTell).ItemStartIndex + lValue + lMax
            
                    If Y - tempY < Bars(BarTell).Item(Tell).Height And Y - tempY > 0 And Bars(BarTell).Item(Tell).CanClick And Bars(BarTell).Item(Tell).ClassMember <> ItemType_ControlPlaceHolder Then
                        
                        ' Select this object
                        SelectSingleItem BarTell, Tell
                        
                        ' Return the selected object
                        Set FindSelectedObject = Bars(BarTell).Item(Tell)
                        
                        Exit Function
                    End If
                
                    tempY = tempY + Bars(BarTell).Item(Tell).Height + BarTitle_ItemSpace
                    
                Next
            
            End If
            
        Next
        
    End If
    
    DeselectAllItems

End Function

Private Function GetSelectedItem(X As Single, Y As Single)

    Dim Tell As Long, tempY As Long, lBarCount As Long
    
    lBarCount = BarCount
    
    If X > BarTitle_Left And X < BarTitle_Left + BarTitle_Width And lBarCount >= 0 Then
    
        tempY = BarTitle_BeginTop
    
        For Tell = BarStartIndex To lBarCount
        
            If Y - tempY < bartitle_height And Y - tempY > 0 Then
                GetSelectedItem = Tell
                Exit For
            Else
            
                If Tell = lBarCount Then
                    GetSelectedItem = -1
                End If
                
            End If
        
            tempY = tempY + bartitle_height + BarTitle_Space + IIf(Bars(Tell).State = BarState_Expanded, Bars(Tell).Height, 0)
            
        Next
    
    Else
        GetSelectedItem = -1
    End If

End Function

Public Function BarPositionY(lIndex As Long) As Long

    Dim Tell As Long
    
    For Tell = BarStartIndex To BarCount
        
        If Tell >= lIndex Then
            Exit For
        End If
        
        BarPositionY = BarPositionY + bartitle_height + BarTitle_Space + IIf(Bars(Tell).State = BarState_Expanded, Bars(Tell).Height, 0)
    Next

End Function

Public Function TextHeight(Str As String) As Single

    TextHeight = UserControl.TextHeight(Str)

End Function

Public Function TextWidth(Str As String) As Single
    
    TextWidth = UserControl.TextWidth(Str)

End Function

Public Function ScaleX(Width As Single, Optional FromScale, Optional ToScale) As Single

    ScaleX = UserControl.ScaleX(Width, FromScale, ToScale)

End Function

Public Function ScaleY(Height As Single, Optional FromScale, Optional ToScale) As Single

    ScaleY = UserControl.ScaleY(Height, FromScale, ToScale)

End Function

Private Sub UserControl_Initialize()

    ' No selected elements
    SelectedItem = -1

End Sub

Private Function GetScrollIndex() As Long

    Dim Tell As Long
    
    For Tell = 0 To vscOverflow.Count - 1
        If vscOverflow(Tell).Visible = False Then
            GetScrollIndex = Tell
            Exit Function
        End If
    Next
    
    GetScrollIndex = vscOverflow.Count
    Load vscOverflow(GetScrollIndex)

End Function

Private Sub vscOverflow_Change(Index As Integer)

    Dim Tell As Long
    
    For Tell = BarStartIndex To BarCount
        If Bars(Tell).ScrollControl Is vscOverflow(Index) Then
            Bars(Tell).DeselectItems
            Bars(Tell).ScrollValue = vscOverflow(Index).Value
        End If
    Next
    
    Redraw

End Sub
