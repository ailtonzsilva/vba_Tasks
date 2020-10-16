VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
    strAppPath As String
    strFileName As String
    strFilePath As String
    strLinkRemote As String
End Type
Private This As ClassType

Property Get strAppPath() As String:                   strAppPath = This.strAppPath:          End Property
Property Let strAppPath(pValue As String):             This.strAppPath = pValue:              End Property
Property Get strFileName() As String:                   strFileName = This.strFileName:       End Property
Property Let strFileName(pValue As String):             This.strFileName = pValue:            End Property
Property Get strFilePath() As String:                   strFilePath = This.strFilePath:       End Property
Property Let strFilePath(pValue As String):             This.strFilePath = pValue:            End Property
Property Get strLinkRemote() As String:                   strLinkRemote = This.strLinkRemote:       End Property
Property Let strLinkRemote(pValue As String):             This.strLinkRemote = pValue:            End Property

Public Function Tips(ByRef pCol As Collection)

    '' Explorer
    pCol.add ":Vex | Explorer"
    
    '' Split
    pCol.add ":sp | Split"
    pCol.add ":vsp | Split vertical"
    
    '' file
    pCol.add ":w | Save"
    pCol.add ":q | Quit"
    pCol.add ":enew | New"
    
    '' Delete lines
    pCol.add ":1,10d | Delete [1-10]"
    
    '' History
    pCol.add ":hist or q: | History"
    
    pCol.add ":[line] put = $MYVIMRC"
    pCol.add ":ol"
    
    pCol.add ":mkview"
    pCol.add ":mkvimrc"
    
    '' macro
    pCol.add "qa...q | Record macro"
    pCol.add "@a | Execute macro"
    
    '' comment
    pCol.add ":2,4s/^/# | comment[1]"
    pCol.add ":'<,'>s/^/# | Comment[2]"
    pCol.add ":'<,'>s/^#/ | UnComment"
        
    '' Lines
    pCol.add ":sort u | Removing duplicate rows in vi"
    pCol.add ":g/^$/d | delete blank linesv"
    pCol.add "yy | Copia a linha atual"
        
    '' Replace
    pCol.add ":%s/foo/bar/g | Encontra cada ocorrência de ‘foo’ (em todas as linhas) e substitui por ‘bar’:"
    pCol.add ":%s/\s\+$//e | Remove unwanted spaces"
    
    '' Change caracters
    pCol.add "~    : Changes the case of current character"
    pCol.add "guu  : Change current line from upper to lower."
    pCol.add "gUU  : Change current LINE from lower to upper."
    pCol.add "guw  : Change to end of current WORD from upper to lower."
    pCol.add "guaw : Change all of current WORD to lower."
    pCol.add "gUw  : Change to end of current WORD from lower to upper."
    pCol.add "gUaw : Change all of current WORD to upper."
    pCol.add "g~~  : Invert case to entire line"
    pCol.add "g~w  : Invert case to current WORD"
    pCol.add "guG : Change to lowercase until the end of document."
        
    If strFileName = "" Then This.strFileName = "Vim_Tips.txt"

End Function

