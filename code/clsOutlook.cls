VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
    strTo As String
    strCC As String
    strBCC As String
    strSubject As String
    strCategory As String
    strBody As String
    strHtmlBody As String
    strStart As String
    strFileNames As String
    strEnd As String
End Type
Private This As ClassType

Property Get strTo() As String:                     strTo = This.strTo:                    End Property
Property Let strTo(pValue As String):               This.strTo = pValue:                   End Property

Property Get strCC() As String:                     strCC = This.strCC:                    End Property
Property Let strCC(pValue As String):               This.strCC = pValue:                   End Property

Property Get strBCC() As String:                    strBCC = This.strBCC:                  End Property
Property Let strBCC(pValue As String):              This.strBCC = pValue:                  End Property

Property Get strSubject() As String:                strSubject = This.strSubject:          End Property
Property Let strSubject(pValue As String):          This.strSubject = pValue:              End Property

Property Get strCategory() As String:               strCategory = This.strCategory:        End Property
Property Let strCategory(pValue As String):         This.strCategory = pValue:             End Property

Property Get strBody() As String:                   strBody = This.strBody:                End Property
Property Let strBody(pValue As String):             This.strBody = pValue:                 End Property

Property Get strHtmlBody() As String:               strHtmlBody = This.strHtmlBody:        End Property
Property Let strHtmlBody(pValue As String):         This.strHtmlBody = pValue:             End Property

Property Get strFileNames() As Variant:             strFileNames = This.strFileNames:      End Property
Property Let strFileNames(pValue As Variant):       This.strFileNames = pValue:            End Property

Property Get strStart() As String:                  strStart = This.strStart:              End Property
Property Let strStart(pValue As String):            This.strStart = pValue:                End Property

Property Get strEnd() As String:                    strEnd = This.strEnd:                  End Property
Property Let strEnd(pValue As String):              This.strEnd = pValue:                  End Property

Function reportAll()
On Error GoTo reportAll_err

Dim OutApp As Object
Dim myItem As Object
Const olAppointmentItem As Long = 1

Set OutApp = CreateObject("Outlook.Application")
Set myItem = OutApp.createItem(olAppointmentItem)

With myItem
    
    .Subject = strSubject
    .body = strBody
    .Start = strStart
    .End = strEnd
    .AllDayEvent = False
    .Save
    
End With

reportAll_Fim:
    Set myItem = Nothing
    Set OutApp = Nothing

    Exit Function
reportAll_err:
    MsgBox err.Description
    Resume reportAll_Fim
    
End Function

Function EnviarEmail()
On Error GoTo Enviar_err

Dim OutApp As Object
Dim OutMail As Object
Dim i As Long

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.createItem(0)

    On Error Resume Next
    With OutMail
        .To = strTo
        .CC = strCC
        .BCC = strBCC
        
        ' Subject
        If (Len(strCategory) > 0) Then
            .Subject = "[" & UCase(strCategory) & "] " & strSubject
        Else
            .Subject = strSubject
        End If
        
        ' Body
        If (Len(strHtmlBody) > 0) Then
            .htmlbody = strHtmlBody
        Else
            .body = strBody
        End If
        
        ' ADICIONAR UM OU MAIS ARQUIVOS EM ANEXO
        If IsArray(strFileNames) Then
            For i = LBound(strFileNames) To UBound(strFileNames)
                .Attachments.add strFileNames(i)
            Next i
        End If
        
        .sEnd
        DoEvents
    End With
    On Error GoTo 0

Enviar_Fim:
    Set OutMail = Nothing
    Set OutApp = Nothing

    Exit Function
Enviar_err:
    MsgBox err.Description
    Resume Enviar_Fim

End Function
