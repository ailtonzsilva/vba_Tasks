VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGit"
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

    pCol.add "#Principal"
    pCol.add "git init  | Iniciar um repositorio local"
    pCol.add "git remote add origin _LinkRepositorio_   | Linkar local com o remoto"
    pCol.add "git pull origin master        | Merge - Sincronizar repositorio local e remoto"
    pCol.add "git add .         | Adicionar arquivos alterados"
    pCol.add "git commit -m ""...""       | Comentar alterações"
    pCol.add "git push origin master        | Subir alterações"
    pCol.add "git branch _NameBranch_       | Nova Branch"
    pCol.add "git branch -a         | Listar branchs"
    pCol.add "git checkout _NameBranch_ | Alterar entre ramificações"
    pCol.add ""
    pCol.add "#Estudos "
    pCol.add "git branch -a         | Listar branchs"
    pCol.add "git stash         | Remover trabalho e guarda em repositorio local"
    pCol.add "git log               | Log"
    pCol.add "git show _hash_           | Mostrar log"
    pCol.add "git diff          | Mostrar ultimas modificações"
    pCol.add "git reset --soft Head~1       | CUIDADO!!! - MOVE O PONTEIRO PARA OUTRO COMMIT"
    
    If strFileName = "" Then This.strFileName = "git_Tips.txt"

End Function


Public Function git_local_remote(ByRef pCol As Collection)

    pCol.add "git remote add origin " & strLinkRemote
    
    If strFileName = "" Then This.strFileName = "git_local_remote.txt"

End Function
