VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XmlHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function FirstElementByXpath(ParentElement As Object, xpath As String) As IXMLDOMElement
    Dim Element As Object
    Set FirstElementByXpath = ParentElement.SelectSingleNode(xpath)
End Function

Public Function FirstElementTextByXpath(ParentElement As Object, xpath As String) As String
    Dim Element As IXMLDOMElement
    Set Element = ParentElement.SelectSingleNode(xpath)
    If Not Element Is Nothing Then
        FirstElementTextByXpath = Element.Text
    End If
End Function

Public Function HasChildElementByXpath(ParentElement, xpath As String)
    HasChildElementByXpath = False
    Dim Element As Object
    Set Element = ParentElement.SelectSingleNode(xpath)
    If Not Element Is Nothing Then
        HasChildElementByXpath = True
    End If
End Function

'
'Public Function ElementsByXpath(ParentElement, xpath As String)
'    Dim Element As Object
'    Set Element = ParentElement.SelectSingleNode(xpath)
'    If Not Element Is Nothing Then
'        FirstElementByXpath = Element
'    End If
'End Function
'
'
'
'Public Function GetElementTextByXpath(ParentElement, xpath As String)
'    Dim Element As Object
'    Set Element = ParentElement.SelectSingleNode(xpath)
'    If Not Element Is Nothing Then
'        GetTextByElementName = Element.Text
'    End If
'End Function
'
'
'Public Function HasChildElementByXpath(ParentElement, xpath As String)
'    HasChildElementWithName = False
'    Dim Element As Object
'    Set Element = ParentElement.SelectSingleNode(xpath)
'    If Not Element Is Nothing Then
'        HasChildElementWithName = True
'    End If
'End Function
'
'Public Function GetAttributeInElementName(ParentElement, xpath As String, attributeName As String)
'    Dim Element As Object
'    Set Element = ParentElement.SelectSingleNode(xpath)
'    If Not Element Is Nothing Then
'        GetAttributeByElementName = Element.getAttribute(attributeName)
'    End If
'End Function
'
'
