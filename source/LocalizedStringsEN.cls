VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrEnglishUS (1033)
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
    Set Strings = New Scripting.Dictionary
    With Strings
        
        .Add "Common.ErrNoDocument", "No active document"
        
        .Add "MainView.frForm", "Formation of compositions"
        .Add "MainView.obByCount", "by columns / rows"
        .Add "MainView.lbColumns", "columns"
        .Add "MainView.lbRows", "rows"
        .Add "MainView.opByArea", "by size"
        .Add "MainView.lbWidth", "mm width"
        .Add "MainView.lbHeight", "mm height"
        .Add "MainView.opByPage", "by page size"
        
        .Add "MainView.frCompose", "Arrangement of compositions"
        .Add "MainView.opOnePage", "all on one page"
        .Add "MainView.opMultiplePages", "each on own page"
        
        .Add "MainView.frSpaces", "Spaces between elements"
        .Add "MainView.lbHorizontalSpace", "mm by width"
        .Add "MainView.lbVerticalSpace", "mm by height"
        
        .Add "MainView.cbNewDocument", "Create new document"
        .Add "MainView.cbGroup", "Group compositions"
        
        .Add "MainView.btnOk", "Start"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
