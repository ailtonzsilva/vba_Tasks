Attribute VB_Name = "mdlMain"
'' #Links
'' [ imageMSO ]
'' https://bert-toolkit.com/imagemso-list.html
'' [ git ]
'' https://githowto.com/pt-BR/create_a_project

'' #customUI
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
'<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
'    <ribbon startFromScratch="false">
'        <tabs>
'            <tab id="CT1" label="Auxiliar" insertBeforeMso="TabHome">
'                <group id="Grp1" label="Ferramentas">
'                    <button id="Btn100" label="Copiar Guia" size="large" onAction="create_fileNew" imageMso="Copy"/>
'                </group>
'                <group id="Grp2" label="Aplicativos">
'                    <button id="Btn200" label="Firefox" size="large" onAction="main_appFirefox" imageMso="MicrosoftOnTheWeb01"/>
'                    <button id="Btn210" label="Freemind" size="large" onAction="main_appFreemind" imageMso="AdpDiagramRelationships"/>
'                    <button id="Btn220" label="Vim" size="large" onAction="main_appGVim" imageMso="CodeEdit"/>
'                    <button id="Btn230" label="Postman" size="large" onAction="main_appPostman" imageMso="ViewsFormView"/>
'                    <button id="Btn240" label="sqlDeveloper" size="large" onAction="main_appSql" imageMso="ViewsAdpDiagramSqlView"/>
'                </group>
'                <group id="Grp3" label="Scripts">
'                    <button id="Btn300" label="Mouse" size="large" onAction="main_scrMouse" imageMso="Fish"/>
'                    <button id="Btn310" label="Tema" size="large" onAction="main_scrTema" imageMso="NewNote"/>
'                    <button id="Btn320" label="Links" size="large" onAction="main_scrLinks" imageMso="FileUpdate"/>
'                    <button id="Btn330" label="Contatos" size="large" onAction="main_scrContatos" imageMso="FileUpdate"/>
'                    <button id="Btn340" label="Criar Txt" size="large" onAction="create_fileTxt" imageMso="FileNew"/>
'                </group>
'
'            </tab>
'        </tabs>
'    </ribbon>
'</customUI>


Option Explicit

Private Sub create_fileTxt(ByVal control As IRibbonControl): tools_createFileTxt: End Sub

Private Sub create_fileNew(ByVal control As IRibbonControl): tools_createNewFile: End Sub

'' APLICATIVOS

Private Sub main_appTema(ByVal control As IRibbonControl): create_runTema: End Sub

Private Sub main_appFirefox(ByVal control As IRibbonControl): run_FirefoxPortable: End Sub

Private Sub main_appFreemind(ByVal control As IRibbonControl): run_freemind: End Sub

Private Sub main_appGVim(ByVal control As IRibbonControl): run_gVimPortable: End Sub

Private Sub main_appPostman(ByVal control As IRibbonControl): run_postman: End Sub

Private Sub main_appSql(ByVal control As IRibbonControl): run_Sqldeveloper: End Sub

Private Sub main_appPython(ByVal control As IRibbonControl): run_Python: End Sub

Private Sub main_appVLCPortable(ByVal control As IRibbonControl): run_VLCPortable: End Sub

'' SCRIPT

Private Sub main_scrMouse(ByVal control As IRibbonControl): create_runMouse: End Sub

Private Sub main_scrTema(ByVal control As IRibbonControl): create_runTema: End Sub

Private Sub main_scrLinks(ByVal control As IRibbonControl): tools_jsonLinks: End Sub

Private Sub main_scrContatos(ByVal control As IRibbonControl): tools_jsonContatos: End Sub

Private Sub main_scrRunUrls(ByVal control As IRibbonControl): tools_createRunUrls: End Sub

Sub create_runUrls(): tools_createRunUrls: End Sub

Sub create_runTema(): tools_createRunTema: End Sub

Sub create_runMouse(): tools_runMouse: End Sub

'Sub create_baseOcultar()
'    tools_baseOcultar
'End Sub




