VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarItem"
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

Enum ItemType
    ItemType_Link = 0                'The default. An item which highlights like a Hyperlink whilst the mouse is over, and can be clicked.
    ItemType_ControlPlaceHolder = 1  'A textual label. This type cannot be highlighted or clicked.
    ItemType_Object = 2              'A placeholder for another ActiveX control. The control is set using the Control property.
End Enum

Private m_lIndex As Long
Private m_sText As String
Private m_sTag As String
Private m_bBold As Boolean
Private m_vControl As Object
Private m_lParent As Object
Private m_lItemData As Variant
Private m_bCanClick As Boolean
Private m_lSpacingAfter As Long
Private m_lIconHandle As StdPicture
Private m_sToolTipText As String
Private m_lClassMember As ItemType
Private m_lTextColor As OLE_COLOR
Private m_lTextColorOver As OLE_COLOR
Private m_bInvokeEvents As Boolean
Private m_bSelected As Boolean

Public Property Get Text() As String

    Text = m_sText
    
End Property

Public Property Let Text(ByVal sText As String)

    If m_sText <> sText Then
        m_sText = sText
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "Text"
    End If
    
End Property

Public Property Get Tag() As String

    Tag = m_sTag
    
End Property

Public Property Let Tag(ByVal sTag As String)

    If m_sTag <> sTag Then
        m_sTag = sTag
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "Tag"
    End If
    
End Property

Public Property Get Bold() As Boolean

    Bold = m_bBold
    
End Property

Public Property Let Bold(ByVal bBold As Boolean)

    If m_bBold <> bBold Then
        m_bBold = bBold
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "Bold"
    End If
    
End Property

Public Property Get Control() As Object

    Set Control = m_vControl
    
End Property

Public Property Let Control(ByVal vControl As Object)

    If Not m_vControl Is vControl Then
        Set m_vControl = vControl
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "Control"
    End If
    
End Property

Public Property Get ItemData() As Variant

    ItemData = m_lItemData
    
End Property

Public Property Let ItemData(ByVal lItemData As Variant)

    If m_lItemData <> lItemData Then
        m_lItemData = lItemData
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "ItemData"
    End If
    
End Property

Public Property Get CanClick() As Boolean

    CanClick = m_bCanClick
    
End Property

Public Property Let CanClick(ByVal bCanClick As Boolean)

    If m_bCanClick <> bCanClick Then
        m_bCanClick = bCanClick
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "CanClick"
    End If
    
End Property

Public Property Get SpacingAfter() As Long

    SpacingAfter = m_lSpacingAfter
    
End Property

Public Property Let SpacingAfter(ByVal lSpacingAfter As Long)

    If m_lSpacingAfter <> lSpacingAfter Then
        m_lSpacingAfter = lSpacingAfter
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "SpacingAfter"
    End If
    
End Property

Public Property Get IconHandle() As StdPicture

    Set IconHandle = m_lIconHandle
    
End Property

Public Property Let IconHandle(ByVal lIconHandle As StdPicture)

    If Not m_lIconHandle Is lIconHandle Then
        Set m_lIconHandle = lIconHandle
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "IconHandle"
    End If
    
End Property

Public Property Get ToolTipText() As String

    ToolTipText = m_sToolTipText
    
End Property

Public Property Let ToolTipText(ByVal sToolTipText As String)

    If m_sToolTipText <> sToolTipText Then
        m_sToolTipText = sToolTipText
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "ToolTipText"
    End If
    
End Property

Public Property Get TextColor() As OLE_COLOR

    TextColor = m_lTextColor
    
End Property

Public Property Let TextColor(ByVal lTextColor As OLE_COLOR)

    If m_lTextColor <> lTextColor Then
        m_lTextColor = lTextColor
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "TextColor"
    End If
    
End Property

Public Property Get TextColorOver() As OLE_COLOR

    TextColorOver = m_lTextColorOver
    
End Property

Public Property Let TextColorOver(ByVal lTextColorOver As OLE_COLOR)

    If m_lTextColorOver <> lTextColorOver Then
        m_lTextColorOver = lTextColorOver
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "TextColorOver"
    End If
    
End Property

Public Property Get Parent() As Object

    Set Parent = m_lParent
    
End Property

Public Property Let Parent(ByVal lParent As Object)

    If Not m_lParent Is lParent Then
        Set m_lParent = lParent
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "Parent"
    End If
    
End Property

Public Property Get ClassMember() As ItemType

    ClassMember = m_lClassMember
    
End Property

Public Property Let ClassMember(ByVal lClassMember As ItemType)

    If m_lClassMember <> lClassMember Then
        m_lClassMember = lClassMember
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "ClassMember"
    End If
    
End Property

Public Property Get Selected() As Boolean

    Selected = m_bSelected
    
End Property

Public Property Let Selected(ByVal bSelected As Boolean)

    If m_bSelected <> bSelected Then
        m_bSelected = bSelected
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "Selected"
    End If
    
End Property

Public Property Get Index() As Long

    Index = m_lIndex
    
End Property

Public Property Let Index(ByVal lIndex As Long)

    If m_lIndex <> lIndex Then
        m_lIndex = lIndex
        If m_bInvokeEvents Then m_lParent.Parent.HandleEvent BarEvent_ItemChanged, Me, "Index"
    End If
    
End Property

Public Property Get InvokeEvents() As Boolean

    InvokeEvents = m_bInvokeEvents
    
End Property

Public Property Let InvokeEvents(ByVal bInvokeEvents As Boolean)

    m_bInvokeEvents = bInvokeEvents
    
End Property

Public Property Get Height() As Long
    
    Dim Tmp As Long, IconHeight As Long, lWidth As Long
    
    Select Case m_lClassMember
    Case ItemType_Object
        
        If Not Control Is Nothing Then
        
            ' A MDIForm hasn't got the property "ScaleMode", so here it will always be twips
            If TypeOf Control.Parent Is MDIForm Then
                ' Only pictureboxes can be placed as "top" childrens (almost), so we must check what type the object is
                If TypeOf Control Is PictureBox Then
                    Tmp = m_lParent.Parent.ScaleY(Control.Height, vbTwips, vbPixels)
                Else
                    ' If not, then just assume it's Pixels, since we cannot really find it out due to the language.
                    Tmp = Control.Height
                End If
            Else
                Tmp = m_lParent.Parent.ScaleY(Control.Height, Control.Parent.ScaleMode, vbPixels)
                'Debug.Print Control.Height
                'Debug.Print Control.Name
            End If
            
        End If

    Case Else
    
        ' Get the size of a normal A
        Tmp = m_lParent.Parent.TextHeight("A")
                
        ' It's faster to access a variable than a property
        lWidth = Width
        
        ' Then, assure that we don't have more lines
        If lWidth > BarTitle_Width - BarTitle_ItemLeft - m_lSpacingAfter Then
            Tmp = Tmp * (Int(lWidth / (BarTitle_Width + BarTitle_Left - m_lSpacingAfter)) + 1)
        End If
    
        ' Make sure that the icon isn't bigger than the text
        If Not m_lIconHandle Is Nothing Then
        
            IconHeight = m_lParent.Parent.ScaleY(m_lIconHandle.Height, vbHimetric, vbPixels)
        
            If IconHeight > Tmp Then
                Tmp = IconHeight
            End If
        
        End If
    
    End Select
    
    ' Return the height
    Height = Tmp
    
End Property

Public Property Get Width() As Long
    
    Dim Tmp As Long
    
    Select Case m_lClassMember
    Case ItemType_Object
        
        If Not Control Is Nothing Then
        
            ' A MDIForm hasn't got the property "ScaleMode", so here it will always be twips
            If TypeOf Control.Parent Is MDIForm Then
                Tmp = m_lParent.Parent.ScaleX(Control.Width, vbTwips, vbPixels)
            Else
                Tmp = m_lParent.Parent.ScaleX(Control.Width, Control.Parent.ScaleMode, vbPixels)
            End If
            
        End If

    Case Else
    
        ' Get the width of the text
        Tmp = m_lParent.Parent.TextWidth(m_sText)
    
    End Select
    
    ' Return the height
    Width = Tmp
    
End Property

Private Sub Class_Initialize()

    ' Default settings
    m_lTextColorOver = RGB(129, 180, 255)
    m_lTextColor = vbBlack
    m_bCanClick = True

End Sub
