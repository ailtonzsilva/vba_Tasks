VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPython"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
    strAppPath As String
    strFileName As String
    strFilePath As String
End Type
Private This As ClassType

Property Get strAppPath() As String:                   strAppPath = This.strAppPath:          End Property
Property Let strAppPath(pValue As String):             This.strAppPath = pValue:              End Property
Property Get strFileName() As String:                   strFileName = This.strFileName:       End Property
Property Let strFileName(pValue As String):             This.strFileName = pValue:            End Property
Property Get strFilePath() As String:                   strFilePath = This.strFilePath:       End Property
Property Let strFilePath(pValue As String):             This.strFilePath = pValue:            End Property

Public Function parseJson(ByRef pCol As Collection)
    
    pCol.add "import json"
    pCol.add ""
    pCol.add "with open('links.json') as jsonfile:"
    pCol.add "    parsed = json.load(jsonfile)"
    pCol.add ""
    pCol.add "print(json.dumps(parsed, indent=2, sort_keys=True))"
    
    If strFileName = "" Then This.strFileName = "parseJson.py"

End Function

Public Function parseXml(ByRef pCol As Collection)

    pCol.add "import xml.dom.minidom"
    pCol.add ""
    pCol.add "uglyxml = ''"
    pCol.add ""
    pCol.add "xml = xml.dom.minidom.parseString(uglyxml)"
    pCol.add ""
    pCol.add "xml_pretty_str = xml.toprettyxml()"
    pCol.add ""
    pCol.add "print(xml_pretty_str)"
    
    If strFileName = "" Then This.strFileName = "parseXml.py"

End Function
