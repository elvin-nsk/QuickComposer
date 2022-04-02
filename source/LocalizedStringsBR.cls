VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrBrazilianPortuguese (1046) by @corelnaveia 31/03/2022
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
  Set Strings = New Scripting.Dictionary
  With Strings
    
    .Add "Common.ErrNoDocument", "Não há documento aberto!"
        
         .Add "MainView.frForm", "Formatação do agrupamento"
         .Add "MainView.obByCount", "Por colunas/linhas"
         .Add "MainView.lbColumns", "Colunas"
         .Add "MainView.lbRows", "Linhas"
         .Add "MainView.opByArea", "Por tamanho predefinido"
         .Add "MainView.lbWidth", "mm largura"
         .Add "MainView.lbHeight", "mm de altura"
         .Add "MainView.opByPage", "Pelo tamanho da página atual"
        
         .Add "MainView.frCompose", "Organizar grupos"
         .Add "MainView.opOnePage", "Em página única"
         .Add "MainView.opMultiplePages", "Em página individual"
        
         .Add "MainView.frSpaces", "Espacejamento"
         .Add "MainView.lbHorizontalSpace", "mm por largura"
         .Add "MainView.lbVerticalSpace", "mm por altura"
        
         .Add "MainView.cbNewDocument", "Organizar em novo documento"
         .Add "MainView.cbGroup", "Organizar e agrupar"
        
         .Add "MainView.btnOk", "Iniciar"

  End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
  Set ILocalizedStrings_Strings = Strings
End Property

