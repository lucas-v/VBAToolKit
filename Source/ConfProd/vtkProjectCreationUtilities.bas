Attribute VB_Name = "vtkProjectCreationUtilities"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeVbaUnitNamesAndPathes
' Author    : Abdelfattah Lahbib
' Date      : 09/05/2013
' Purpose   : - Initialize DEV project ConfSheet with vbaunit module names and pathes
'             - Return True if module names and paths are initialized without error
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeVbaUnitNamesAndPathes(project As String) As Boolean
    Dim tableofvbaunitname(17) As String
        tableofvbaunitname(0) = "VbaUnitMain"
        tableofvbaunitname(1) = "Assert"
        tableofvbaunitname(2) = "AutoGen"
        tableofvbaunitname(3) = "IAssert"
        tableofvbaunitname(4) = "IResultUser"
        tableofvbaunitname(5) = "IRunManager"
        tableofvbaunitname(6) = "ITest"
        tableofvbaunitname(7) = "ITestCase"
        tableofvbaunitname(8) = "ITestManager"
        tableofvbaunitname(9) = "RunManager"
        tableofvbaunitname(10) = "TestCaseManager"
        tableofvbaunitname(11) = "TestClassLister"
        tableofvbaunitname(12) = "TesterTemplate"
        tableofvbaunitname(13) = "TestFailure"
        tableofvbaunitname(14) = "TestResult"
        tableofvbaunitname(15) = "TestRunner"
        tableofvbaunitname(16) = "TestSuite"
        tableofvbaunitname(17) = "TestSuiteManager"
    Dim i As Integer, cm As vtkConfigurationManager, ret As Boolean, nm As Integer, nc As Integer, ext As String
    Set cm = vtkConfigurationManagerForProject(project)
    nc = cm.getConfigurationNumber(vtkProjectForName(project).projectDEVName)
    ret = (nc > 0)
    For i = LBound(tableofvbaunitname) To UBound(tableofvbaunitname)
        nm = cm.addModule(tableofvbaunitname(i))
        ret = ret And (nm > 0)
        If i <= 0 Then      ' It's a Standard Module (WARNING, magical number)
            ext = ".bas"
           Else
            ext = ".cls"    ' It's a Class Module
        End If
        cm.setModulePathWithNumber path:="Source\VbaUnit\" & tableofvbaunitname(i) & ext, numModule:=nm, numConfiguration:=nc
    Next i
    vtkInitializeVbaUnitNamesAndPathes = ret
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkActivateReferences
' Author    : Abdelfattah Lahbib
' Date      : 26/04/2013
' Purpose   : - Check that workbook is open
'             - Activate VBIDE and +-scripting references
'---------------------------------------------------------------------------------------
Public Function vtkActivateReferences(workbookName As String)
    If VtkWorkbookIsOpen(workbookName) = True Then     'if the workbook is ope
        On Error Resume Next ' if the first extention is already activated, we will try to activate the second one
        Workbooks(workbookName).VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 0, 0  ' +- to activate Scripting : Microsoft scripting runtime
        Workbooks(workbookName).VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 0, 0 ' to activate VBIDE: Microsoft visual basic for applications extensibility 5.3
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkCreateProjectRootFolder
' Author    : Lucas Vitorino
' Purpose   : - Create a folder with a given name in a given path
'             - Return string : "$path\$name"
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateProjectRootFolder(path As String, name As String)
        MkDir path & "\" & name
        vtkCreateProjectRootFolder = path & "\" & name
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkCreateProjectFolderTree
' Author    : Lucas Vitorino
' Purpose   : Create the arborescence of folders of a VTK project in a given directory
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateProjectFolderTree(rootPath As String)
    ' Create Delivery folder
    MkDir rootPath & "\" & "Delivery"
    ' Create Project folder
    MkDir rootPath & "\" & "Project"
    ' Create Tests folder
    MkDir rootPath & "\" & "Tests"
    ' Create GitLog Folder
    MkDir rootPath & "\" & "GitLog"
    ' Create Source folder
    MkDir rootPath & "\" & "Source"
    ' Create ConfProd folder
    MkDir rootPath & "\" & "Source" & "\" & "ConfProd"
    ' Create ConfTest folder
    MkDir rootPath & "\" & "Source" & "\" & "ConfTest"
    ' Create VbaUnit folder
    MkDir rootPath & "\" & "Source" & "\" & "VbaUnit"
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkCreateProjectDEVWorkbook
' Author    : Lucas Vitorino
' Purpose   : Create the DEV Excel Workbook of a VTK project.
' Notes     : - The workbook is not closed after its creation.
'             - TODO : allow choice of file format ?
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateProjectDEVWorkbook(project As vtkProject, rootPath As String)
    'Save created project with xlsm extention
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectDEVStandardRelativePath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
    'Rename Project
    Workbooks(project.workbookDEVName).VBProject.name = project.projectDEVName
    'call function to activate references
    vtkActivateReferences (project.workbookDEVName)
    'initialize configuration Sheet with VBAUnit modules
    vtkInitializeVbaUnitNamesAndPathes project:=project.projectName
    ' Save Development Project Workbook
    Workbooks(project.workbookDEVName).Save
    ' TODO : is the following line really useful ?
    'Workbooks(project.workbookDEVName).Activate
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkCreateDeliveryWorkbook
' Author    : Lucas Vitorino
' Purpose   : Create the delivery Excel Workbook of a VTK project
' Notes     : - The workbook is closed after its creation.
'             - TODO : allow choice of file format
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateProjectDeliveryWorkbook(project As vtkProject, rootPath As String)
    'Create delivery workbook
    Workbooks.Add.SaveAs (rootPath & "\" & project.projectStandardRelativePath), FileFormat:=(52) '52 is xlsm format
    'Rename Project
    Workbooks(project.workbookName).VBProject.name = project.projectName
    'call function to activate references
    vtkActivateReferences (project.workbookName)
    ' A module must be added in the Excel File for the project parameters to be saved
    Workbooks(project.workbookName).VBProject.VBComponents.Add ComponentType:=vbext_ct_StdModule
    ' Save and Close Delivery Project WorkBook
    Workbooks(project.workbookName).Close SaveChanges:=True
End Function

