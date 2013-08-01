VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkCreateProjectTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private Sub Class_Initialize()
    Set mManager = New TestCaseManager
End Sub

Private Property Get ITestCase_Manager() As TestCaseManager
    Set ITestCase_Manager = mManager
End Property

Private Property Get ITest_Manager() As ITestManager
    Set ITest_Manager = mManager
End Property

Private Sub ITestCase_SetUp(Assert As IAssert)
    Set mAssert = Assert
End Sub


Private Sub ITestCase_TearDown()
    ' Reset all configuration managers
    vtkResetConfigurationManagers
    ' Make sure to be out of the folder to clean
    Dir (vtkTestPath)
    ' Do not display messagebox to ask to save the project
    Application.DisplayAlerts = False
    On Error Resume Next
    ' Close Created WorkBook
    Workbooks(vtkTestProjectName & ".xlsm").Close
    Workbooks(vtkTestProjectName & "_DEV" & ".xlsm").Close
    ' Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName & "_DEV.xlsm"
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Delivery\" & vtkTestProjectName & ".xlsm"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfProd"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfTest"
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit\*"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\GitLog"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Tests"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Delivery"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Project"
    RmDir vtkTestPath & "\" & vtkTestProjectName
    Application.DisplayAlerts = True
End Sub

Private Sub OpenDeliveryProject()
    Dim path As String
    Dim fso As New FileSystemObject
    path = fso.GetParentFolderName(ActiveWorkbook.path) & "\" & vtkConfigurationManagerForProject(vtkTestProjectName).getConfigurationPath(vtkTestProjectName)
    Workbooks.Open Filename:=path
End Sub

Public Sub TestCreateProjectDoesNothingWhenPathDoesntExist()
    ' The createProject must do nothing when path doesn't exist
    Dim InexistentPath As String
    Dim returnValue As Long
    InexistentPath = vtkTestPath & "\InexistentFolder"
    returnValue = vtkCreateProject(path:=InexistentPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(InexistentPath & "\" & vtkTestProjectName, vbDirectory) = "", "The project main folder must exist"
    mAssert.Equals returnValue, 76, "The project creation must fail"
End Sub

Public Sub TestCreateProjectAlreadyExisting()
    ' The createProject must do nothing when project already exists
    Dim returnValue As Long
    ' Must respect order mkdir before call vtkcreateproject
    MkDir vtkTestPath & "\" & vtkTestProjectName
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The existing project main folder must exist"
    mAssert.Equals returnValue, 75, "The project creation must fail"
End Sub

Public Sub TestFilesAndFoldersAreCreatedCorrectly()
    Dim returnValue As Long
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    ' Test the creation of a new folder named like the project in the path
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName, vbDirectory) <> "", "The project main folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
    ' Test the creation of each subfolder in the tree
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Project", vbDirectory) <> "", "The Project folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Source", vbDirectory) <> "", "The Source folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Tests", vbDirectory) <> "", "The Tests folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfProd", vbDirectory) <> "", "source\ConfProd folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\ConfTest", vbDirectory) <> "", "source\ConfTest folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit", vbDirectory) <> "", "The source\VbaUnit folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\GitLog", vbDirectory) <> "", "The ..\GitLog folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\Delivery", vbDirectory) <> "", "The ..\Delivery folder must exist"
    mAssert.Equals returnValue, 0, "The project creation must succeed"

'    ' Test that the xlsm DEV Excel file is created in the right path
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\" & "Project" & "\" & vtkTestProjectName & "_DEV.xlsm", vbDirectory) <> "", "the DEV workbook must be created"
'    ' Test that xlsm delivery Excel file is created in the right path
    mAssert.Should Dir(vtkTestPath & "\" & vtkTestProjectName & "\" & "Delivery" & "\" & vtkTestProjectName & ".xlsm", vbDirectory) <> "", "the delivery workbook must be created"
'    ' Test that the xlsm DEV Excel file is correctly named
    mAssert.Equals Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.name, vtkTestProjectName & "_DEV", "the DEV workbook name must be " & vtkTestProjectName & "_DEV"
'    ' Test that the DEV project is correctly named
    mAssert.Should Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.name = vtkTestProjectName & "_DEV", "the workbook name must be " & vtkTestProjectName & "_DEV"
'    ' Test that the delivery project is correctly named
    OpenDeliveryProject
    mAssert.Equals Workbooks(vtkTestProjectName & ".xlsm").VBProject.name, vtkTestProjectName, "the delivery workbook name must be " & vtkTestProjectName
    
End Sub

Public Sub TestReferencesOfDevAndDeliveryWorkbooks()
    Dim returnValue As Long
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    ' In the DEV project, this function will count the number of extensions before executing the function ,
    ' and the number of extensions must be equal to (initial number of extensions + number of added extensions)
    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
    OpenDeliveryProject
    For i = 1 To Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.References.Count
        If ((Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.References(i).GUID) = "{420B2830-E718-11CF-893D-00A0C9054228}") Or (Workbooks(vtkTestProjectName & "_DEV.xlsm").VBProject.References(i).GUID = "{0002E157-0000-0000-C000-000000000046}") Then
            j = j + 1
        End If
    Next
    mAssert.Equals j, 2, "the references number of the DEV workbook must be equal to" & i + 2
    
    j = 0
    ' In the delivery project, this function will count the number of extensions before executing the function ,
    ' and the number of extensions must be equal to (initial number of extensions + number of added extensions)
    For i = 1 To Workbooks(vtkTestProjectName & ".xlsm").VBProject.References.Count
        If ((Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID) = "{420B2830-E718-11CF-893D-00A0C9054228}") Or (Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID = "{0002E157-0000-0000-C000-000000000046}") Then
            j = j + 1
        End If
    Next
    mAssert.Equals j, 2, "the references number of the delivery workbook must be equal to" & i + 2
End Sub

'Public Sub TestReferencesOfDeliveryWorkbookWasActivated()
'    Dim returnValue As Long
'    Dim i As Integer
'    Dim j As Integer
'    j = 0
'    ' In the delivery project, this function will count the number of extensions before executing the function ,
'    ' and the number of extensions must be equal to (initial number of extensions + number of added extensions)
'    returnValue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'    OpenDeliveryProject
'
'    For i = 1 To Workbooks(vtkTestProjectName & ".xlsm").VBProject.References.Count
'        If ((Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID) = "{420B2830-E718-11CF-893D-00A0C9054228}") Or (Workbooks(vtkTestProjectName & ".xlsm").VBProject.References(i).GUID = "{0002E157-0000-0000-C000-000000000046}") Then
'            j = j + 1
'        End If
'    Next
'    mAssert.Equals j, 2, "the references number of the delivery workbook must be equal to" & i + 2
'End Sub
'Public Sub TestDevWorkbookPathCorrectlyWritted()
''test that path  of dev workbook is correctly writted on configuration sheet
'    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine - 2), Workbooks(vtkTestProjectName & ".xlsm").FullNameURLEncoded
'End Sub
'Public Sub TestDevWorkbookNameCorrectlyWritted()
''test that name  of dev workbook is correctly writted  on configuration sheet
'    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDevRange & vtkFirstLine - 3), Workbooks(vtkTestProjectName & ".xlsm").name
'End Sub
'Public Sub TestDeliveryWorkbookPathCorrectlyWritted()
''test that path  of delivery workbook is correctly writted on configuration sheet
'    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & vtkFirstLine - 2), Workbooks(vtkTestProjectName & ".xlsm").FullNameURLEncoded
'End Sub
'Public Sub TestDeliveryWorkbookNameCorrectlyWritted()
''test that name  of delivery workbook is correctly writted on configuration sheet
'    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'    mAssert.Equals ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleDeliveryRange & vtkFirstLine - 3), Workbooks(vtkTestProjectName & ".xlsm").name
'End Sub
'Public Sub TestDevWorkbookIsActive()
''test that active workbook is the dev one and not delivery
'    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'    mAssert.Equals ActiveWorkbook.name, vtkTestProjectName & ".xlsm", "the dev workbook must be active"
'End Sub
'Public Sub TestExportVbaUnitModule()
'  'test that exported file was not empty , files was correctly named
'  'the test will compare "vbaunit" file of vbatoolkit and exported file
'  Dim fs As Object
'  Set fs = CreateObject("Scripting.FileSystemObject")
'
'  'call create project function
'  returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'  'vbaunit files of vbatoolkit project
'  sourceVbaUnitFile = Dir(vtkInstallPath & "\source\VbaUnit\*.*", vbDirectory)
'
' 'list of files in the folder
' Do While sourceVbaUnitFile <> ""
'    'ignore these two name
'    If sourceVbaUnitFile <> "." And sourceVbaUnitFile <> ".." Then
'      'test that exported files was not empty
'       mAssert.Should IsEmpty(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit\" & sourceVbaUnitFile) = False, sourceVbaUnitFile & "should not be empty"
'      'test name, extension, and the existence of the file in the created project
'       mAssert.Should fs.FileExists(vtkTestPath & "\" & vtkTestProjectName & "\source\VbaUnit\" & sourceVbaUnitFile) = True, sourceVbaUnitFile & "must exist"
'    End If
'  'point to the next element
'  sourceVbaUnitFile = Dir
' Loop
'
'End Sub
'Public Sub TestImportModules()
''test that modules was imported to the test workbook
'Dim returnvalue As String
'Dim NbrMod1 As Integer
'
'    returnvalue = vtkCreateProject(path:=vtkTestPath, name:=vtkTestProjectName, displayError:=False)
'    NbrMod1 = Workbooks(vtkTestProjectName & ".xlsm").VBProject.VBComponents.Count
''18 vbaunitmodule + 4 default module (4sheets + 1 workbook )
'mAssert.Equals NbrMod1, 18 + 4 + 1, "18 modules must be imported"
'End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateProjectDoesNothingWhenPathDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreateProjectAlreadyExisting"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestFilesAndFoldersAreCreatedCorrectly"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferencesOfDevAndDeliveryWorkbooks"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestCreateProjectDoesNothingWhenPathDoesntExist": TestCreateProjectDoesNothingWhenPathDoesntExist
        Case "TestCreateProjectAlreadyExisting": TestCreateProjectAlreadyExisting
        Case "TestFilesAndFoldersAreCreatedCorrectly": TestFilesAndFoldersAreCreatedCorrectly
        Case "TestReferencesOfDevAndDeliveryWorkbooks": TestReferencesOfDevAndDeliveryWorkbooks
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub

