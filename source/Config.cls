VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Config v2022.02.13
' хранение, загрузка и сохранение настроек
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public OptionComposeByCount As Boolean
Public OptionComposeByArea As Boolean
Public OptionComposeByPage As Boolean
Public OptionOnePage As Boolean
Public OptionMultiplePages As Boolean
Public OptionNewDocument As Boolean
Public OptionGroup As Boolean
Public Columns As Long
Public Rows As Long
Public Width As Double
Public Height As Double
Public HorizontalSpace As Double
Public VerticalSpace As Double

'-------------------------------------------------------------------------------

Private Sub Assign(ByVal Action As iniAction)

    With iniFile.Create
    
        iniFilePrepare .Self, Action
        
        .KeyVar "OptionComposeByCount", OptionComposeByCount, False
        .KeyVar "OptionComposeByArea", OptionComposeByArea, False
        .KeyVar "OptionComposeByPage", OptionComposeByPage, True
        .KeyVar "OptionOnePage", OptionOnePage, False
        .KeyVar "OptionMultiplePages", OptionMultiplePages, True
        .KeyVar "OptionNewDocument", OptionNewDocument, True
        .KeyVar "OptionGroup", OptionGroup, False
        .KeyVar "Columns", Columns, 5&
        .KeyVar "Rows", Rows, 5&
        .KeyVar "Width", Width, 210#
        .KeyVar "Height", Height, 297#
        .KeyVar "HorizontalSpace", HorizontalSpace, 0#
        .KeyVar "VerticalSpace", VerticalSpace, 0#
    
    End With
    
End Sub

'===============================================================================

Public Function CreateAndLoad() As Config
    Set CreateAndLoad = New Config
    CreateAndLoad.Load
End Function

Private Sub Class_Terminate()
    Save
End Sub

'===============================================================================

Public Sub Load()
    Assign iniRead
End Sub

Public Sub Save()
    Assign iniWrite
End Sub

Public Property Get Self() As Config
    Set Self = Me
End Property

'===============================================================================

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

    Dim iniFilePath As String

    With iniFile
        Select Case True
        Case FileExists(GMSManager.UserGMSPath & iniFileName)
            .File = Application.GMSManager.UserGMSPath & iniFileName
            .Action = Action
        Case FileExists(GMSManager.GMSPath & iniFileName)
            .File = Application.GMSManager.GMSPath & iniFileName
            .Action = Action
        Case Else
            .File = GMSManager.UserGMSPath & iniFileName
            If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
        End Select
    End With

End Sub

Public Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Public Sub SaveStrToFile(ByVal Content As String, _
                                                 ByVal File As String)
    Dim FileNum As Long
    FileNum = VBA.FreeFile
    Open File For Output Access Write As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub
