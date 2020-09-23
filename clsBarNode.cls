VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarNode"
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

Private m_sToolTipText As String
Private m_sCaption As String
Private m_lIcon As Long
Private m_lBackColor As Long
Private m_lLineColor As Long
Private m_lMaxItems As Long
Private m_sTitleFont As IFontDisp
Private m_lTitleBackColorLight As Long
Private m_lTitleBackColorDark As Long
Private m_lTitleForeColor As Long
Private m_lTitleForeColorOver As Long
Private m_lButtonBorder As Long
Private m_lButtonFace As Long
Private m_lButtonShadow As Long
Private m_bCanExpand As Boolean
Private m_eState As BarState
Private m_sTag As String
Private m_lItemData As Variant
Private m_lIndex As Long
Private m_bInvokeEvents As Boolean
Private m_ctloParent As uSideBar

' Public variables
Public Height As Long
Public ScrollControl As Object
Public ScrollValue As Long

' Objects
Private Items() As clsBarItem

Public Property Get ItemPositionY(Index As Variant) As Long
    
    Dim lpItem As clsBarItem, tempY As Long, Tell As Long, lMax As Long, lValue As Long
    
    ' Get the item
    Set lpItem = Item(Index)
    If lpItem Is Nothing Then Exit Property
    
    ' Find the Y-position of this bar
    tempY = m_ctloParent.BarPositionY(m_lIndex) + BarTitle_Height + (BarTitle_UpAndDownSpace * 2) + 3
    
    If m_ctloParent.Bar(m_lIndex).ScrollControl Is Nothing Then
        lMax = ItemCount
        lValue = 0
    Else
        lMax = m_ctloParent.Bar(m_lIndex).MaxItems
        lValue = m_ctloParent.Bar(m_lIndex).ScrollControl.Value
    End If
    
    For Tell = ItemStartIndex + lValue To ItemStartIndex + lValue + lMax

        If lpItem.Index = Tell Then
            Exit For
        End If
    
        tempY = tempY + Item(Tell).Height + BarTitle_ItemSpace
        
    Next
    
    ItemPositionY = tempY
    
End Property

Public Property Get ItemCount() As Long

    ItemCount = SafeUBound(VarPtrArray(Items))
    
End Property

Public Property Get ItemStartIndex() As Long

    ItemStartIndex = 0 ' Always 0
    
End Property

Public Property Get ToolTipText() As String

    ToolTipText = m_sToolTipText
    
End Property

Public Property Let ToolTipText(ByVal sToolTipText As String)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, m_sToolTipText
    m_sToolTipText = sToolTipText
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get Caption() As String

    Caption = m_sCaption
    
End Property

Public Property Let Caption(ByVal sCaption As String)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, sCaption
    m_sCaption = sCaption
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get Icon() As Long

    Icon = m_lIcon
    
End Property

Public Property Let Icon(ByVal lIcon As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lIcon
    m_lIcon = lIcon
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get LineColor() As Long
    LineColor = m_lLineColor
End Property

Public Property Let LineColor(ByVal lLineColor As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lLineColor
    m_lLineColor = lLineColor
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get ButtonBorder() As Long

    ButtonBorder = m_lButtonBorder
    
End Property

Public Property Let ButtonBorder(ByVal lButtonBorder As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lButtonBorder
    m_lButtonBorder = lButtonBorder
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get ButtonFace() As Long

    ButtonFace = m_lButtonFace
    
End Property

Public Property Let ButtonFace(ByVal lButtonFace As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lButtonFace
    m_lButtonFace = lButtonFace
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get ButtonShadow() As Long

    ButtonShadow = m_lButtonShadow
    
End Property

Public Property Let ButtonShadow(ByVal lButtonShadow As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lButtonShadow
    m_lButtonShadow = lButtonShadow
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get BackColor() As Long

    BackColor = m_lBackColor
    
End Property

Public Property Let BackColor(ByVal lBackColor As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lBackColor
    m_lBackColor = lBackColor
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get TitleFont() As IFontDisp

    Set TitleFont = m_sTitleFont
    
End Property

Public Property Let TitleFont(ByVal sTitleFont As IFontDisp)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, sTitleFont
    Set m_sTitleFont = sTitleFont
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get TitleForeColor() As Long

    TitleForeColor = m_lTitleForeColor
    
End Property

Public Property Let TitleForeColor(ByVal lTitleForeColor As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lTitleForeColor
    m_lTitleForeColor = lTitleForeColor
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get TitleForeColorOver() As Long

    TitleForeColorOver = m_lTitleForeColorOver
    
End Property

Public Property Let TitleForeColorOver(ByVal lTitleForeColorOver As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lTitleForeColorOver
    m_lTitleForeColorOver = lTitleForeColorOver
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get CanExpand() As Boolean

    CanExpand = m_bCanExpand
    
End Property

Public Property Let CanExpand(ByVal bCanExpand As Boolean)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, bCanExpand
    m_bCanExpand = bCanExpand
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get State() As BarState

    State = m_eState
    
End Property

Public Property Let State(ByVal eState As BarState)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, eState
    m_eState = eState
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get Tag() As String

    Tag = m_sTag
    
End Property

Public Property Let Tag(ByVal sTag As String)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, sTag
    m_sTag = sTag
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get ItemData() As Variant

    ItemData = m_lItemData
    
End Property

Public Property Let ItemData(ByVal lItemData As Variant)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lItemData
    m_lItemData = lItemData
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get Index() As Long

    Index = m_lIndex
    
End Property

Public Property Let Index(ByVal lIndex As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lIndex
    m_lIndex = lIndex
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get TitleBackColorLight() As Long

    TitleBackColorLight = m_lTitleBackColorLight
    
End Property

Public Property Let TitleBackColorLight(ByVal lTitleBackColorLight As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lTitleBackColorLight
    m_lTitleBackColorLight = lTitleBackColorLight
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get TitleBackColorDark() As Long

    TitleBackColorDark = m_lTitleBackColorDark
    
End Property

Public Property Let TitleBackColorDark(ByVal lTitleBackColorDark As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lTitleBackColorDark
    m_lTitleBackColorDark = lTitleBackColorDark
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get MaxItems() As Long

    MaxItems = m_lMaxItems
    
End Property

Public Property Let MaxItems(ByVal lMaxItems As Long)

    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changing, m_lIndex, lMaxItems
    m_lMaxItems = lMaxItems
    If m_bInvokeEvents Then m_ctloParent.HandleEvent BarEvent_Changed, m_lIndex
    
End Property

Public Property Get Parent() As uSideBar

    Set Parent = m_ctloParent
    
End Property

Public Property Set Parent(ctloParent As uSideBar)

    Set m_ctloParent = ctloParent
    
End Property

Public Property Get InvokeEvents() As Boolean

    InvokeEvents = m_bInvokeEvents
    
End Property

Public Property Let InvokeEvents(ByVal bInvokeEvents As Boolean)

    m_bInvokeEvents = bInvokeEvents
    
End Property

Public Property Get Item(Index As Variant) As clsBarItem
    
    Dim Tell As Long
        
    If IsNumeric(Index) Then
        If Index < ItemStartIndex Or Index > ItemCount Then Exit Property
        Set Item = Items(Index)
    Else
    
        For Tell = 0 To ItemCount
        
            If LCase(Index) = LCase(Items(Tell).ItemData) Then
                Set Item = Items(Tell)
                Exit Property
            End If
        
        Next
    
    End If
    
End Property

Private Sub Class_Initialize()

    m_lMaxItems = -1
    m_lLineColor = vbWhite
    m_lBackColor = RGB(214, 223, 247)
    m_lTitleBackColorLight = vbWhite
    m_lTitleBackColorDark = RGB(198, 210, 246)
    m_lTitleForeColor = RGB(82, 129, 211)
    m_lTitleForeColorOver = RGB(129, 180, 255)
    m_lButtonBorder = &HE4BAAC
    m_lButtonShadow = 13281696
    m_lButtonFace = vbWhite

End Sub

Private Sub Class_Terminate()

    ' Clean up
    Set m_ctloParent = Nothing
        
End Sub

Public Sub DeselectItems()
    
    Dim Tell As Long
    
    For Tell = ItemStartIndex To ItemCount
        Items(Tell).Selected = False
    Next

End Sub

Public Sub SelectItem(Index As Long)

    ' Deselect everything else
    DeselectItems
    
    ' Select the item
    Items(Index).Selected = True

End Sub

Public Function AddItem(Optional Key As String, Optional Text As String, Optional ClassMember As ItemType) As clsBarItem

    ' Realllocate the array
    ReDim Preserve Items(SafeUBound(VarPtrArray(Items)) + 1)
    
    ' Create a new bar
    Set Items(UBound(Items)) = New clsBarItem
    
    Items(UBound(Items)).Parent = Me
    Items(UBound(Items)).ItemData = Key
    Items(UBound(Items)).Text = Text
    Items(UBound(Items)).Index = UBound(Items)
    Items(UBound(Items)).ClassMember = ClassMember
    Items(UBound(Items)).InvokeEvents = True
    
    ' Return the created object
    Set AddItem = Items(UBound(Items))
    
    ' Tell about the change
    If Not m_ctloParent Is Nothing Then
        m_ctloParent.HandleEvent BarEvent_ItemAdded, Me.Index, AddItem.Index
    End If

End Function

Public Sub RemoveItem(ByVal Index As Long)

    Dim Tell As Long
    
    If Index < ItemStartIndex Or Index > ItemCount Then
        Exit Sub
    End If
    
    Set Items(Index) = Nothing
    
    For Tell = Index To ItemCount - 1
        CopyMemory ByVal VarPtr(Items(Tell)), ByVal VarPtr(Items(Tell + 1)), 4
    Next
    
    ' Reallocate the array to its right size.
    If UBound(Items) > 0 Then
        ReDim Preserve Items(UBound(Items) - 1)
    Else
        Erase Items
    End If
    
    ' Tell about the change
    If Not m_ctloParent Is Nothing Then
        m_ctloParent.HandleEvent BarEvent_ItemRemoved, Me, Index
    End If

End Sub

Public Property Get ItemExists(ByVal Index As Variant) As Boolean

    Dim Tell As Long
    
    If IsNumeric(Index) Then
    
        If Index < ItemStartIndex Or Index > ItemCount Then
            ItemExists = False
            Exit Property
        End If
        
        If Items(Index) Is Nothing Then
            ItemExists = False
            Exit Property
        End If
    
    Else
    
        For Tell = 0 To ItemCount
        
            If LCase(Index) = LCase(Items(Tell).ItemData) Then
                ItemExists = True
                Exit Property
            End If
        
        Next
    
        Exit Property
    End If
    
    ItemExists = True

End Property
