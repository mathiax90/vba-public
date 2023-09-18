Attribute VB_Name = "XmlHelperTest"
Option Explicit

Dim xml As Object
Dim xpath As String
Dim XmlHelper As New XmlHelperClass
Dim Element As IXMLDOMElement
Dim obj As Object



Sub GetTestXml()
    Set xml = New DOMDocument60
    xml.LoadXML ("<PERSON_LIST><PERSON><ID>0</ID><FAM>Иванов</FAM><IM>Иван</IM><OT>Иванович</OT><DOC TYPE=""PASP"">111111</DOC></PERSON><PERSON><ID>1</ID><FAM>Сидоров</FAM><IM>Сидор</IM><OT>Сидорович</OT><DOC TYPE=""PASP"">222222</DOC></PERSON></PERSON_LIST>")

End Sub

Sub FirstElementByXpathTest()
    Call GetTestXml
    xpath = "PERSON_LIST/PERSON/FAM"
    Set Element = XmlHelper.FirstElementByXpath(xml, xpath)
    If Element.Text = "Иванов" Then
        Debug.Print "FirstElementByXpathTest ok"
    Else
        Debug.Print "FirstElementByXpathTest error"
    End If
End Sub

Sub FirstElementTextByXpathTest()
    Call GetTestXml
    xpath = "PERSON_LIST/PERSON/FAM"
    If XmlHelper.FirstElementTextByXpath(xml, xpath) = "Иванов" Then
        Debug.Print "FirstElementTextByXpathTest ok"
    Else
        Debug.Print "FirstElementTextByXpathTest error"
    End If
End Sub




'Sub HasChildElementWithNameTest()
'    Call GetTestXml
'    xpath = "PERSON_LIST/PERSON/FAM"
'    Debug.Print XmlHelper.HasChildElementWithName(xml, xpath)
'End Sub
'
'Sub GetAttributeByElementNameTest()
'    Call GetTestXml
'    xpath = "PERSON_LIST/PERSON/DOC"
'    Debug.Print XmlHelper.GetAttributeByElementName(xml, xpath, "TYPE")
'End Sub
'
'
'Sub test()
'    Call GetTestXml
'    xpath = "/PERSON_LIST/PERSON"
'    Set PersonElement = xml.SelectSingleNode(xpath)
'    Debug.Print PersonElement.Text
'    Set DocElement = PersonElement.SelectSingleNode("DOC")
'    If DocElement Is Nothing Then
'        Debug.Print "ELement not found"
'        Exit Sub
'    End If
'    Debug.Print DocElement.Text
'End Sub
'
'
