Imports System.Environment
Imports Inventor
Public Class MyButton
    Private ReadOnly _inventor As Application
    Private _buttonDefinition As ButtonDefinition
    Private _control As CommandControl
    Private _ribbonPanel As RibbonPanel
    Public Sub New(inventor As Application)
        _inventor = inventor
        SetupButtonDefinition()
        AddButtonDefinitionToRibbon()
    End Sub
    Public Sub Unload()
        _control = Nothing
        _buttonDefinition.Delete()
        _buttonDefinition = Nothing
        _ribbonPanel = Nothing
    End Sub
    Private Sub SetupButtonDefinition()
        Dim conDefs As ControlDefinitions = _inventor.CommandManager.ControlDefinitions

        _buttonDefinition = Nothing
        _buttonDefinition = conDefs.AddButtonDefinition(
            "Cross-Cutter",
            "CJSC_VB",
            CommandTypesEnum.kEditMaskCmdType,
            Guid.NewGuid().ToString(),
            "CJSC_VB",
            "CJSC_VB")

        AddHandler _buttonDefinition.OnExecute, AddressOf MyButton_OnExecute
        _buttonDefinition.StandardIcon = PictureDispConverter.ToIPictureDisp(My.Resources.CJSC_16)
        _buttonDefinition.LargeIcon = PictureDispConverter.ToIPictureDisp(My.Resources.CJSC_32)
    End Sub
    Private Sub AddButtonDefinitionToRibbon()
        Dim ribbon As Ribbon = _inventor.UserInterfaceManager.Ribbons("Part")
        Dim ribbonTab As RibbonTab = ribbon.RibbonTabs("id_TabModel")

        _ribbonPanel = ribbonTab.RibbonPanels.Add("Modify+", "ModelModifyPlus_Panel", "{FDF39F00-2EF3-45BF-BFE3-709AEA43B660}", "id_PanelP_ModelModify")
        _control = _ribbonPanel.CommandControls.AddButton(_buttonDefinition)
    End Sub
    Private Sub MyButton_OnExecute(Context As NameValueMap)
        Try
            Dim rule As New ThisRule()
            rule.ThisApplication = _inventor
            rule.Main()
        Catch ex As Exception
            MsgBox("Something went wrong while running rule. Message#2: " & ex.Message)
        End Try
    End Sub
End Class