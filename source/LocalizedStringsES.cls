VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrSpanish (1034) by @corelnaveia 31/03/2022
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
  Set Strings = New Scripting.Dictionary
  With Strings
    
    .Add "Common.ErrNoDocument", "No hay documento activo"
        
         .Add "MainView.frForm", "Formación de composiciones"
         .Add "MainView.obByCount", "por columnas/filas"
         .Add "MainView.lbColumns", "columnas"
         .Add "MainView.lbRows", "filas"
         .Add "MainView.opByArea", "por tamaño"
         .Add "MainView.lbWidth", "mm de ancho"
         .Add "MainView.lbHeight", "mm altura"
         .Add "MainView.opByPage", "por tamaño de página"
        
         .Add "MainView.frCompose", "Disposición de composiciones"
         .Add "MainView.opOnePage", "todo en una página"
         .Add "MainView.opMultiplePages", "cada uno en su propia página"
        
         .Add "MainView.frSpaces", "Espacios entre elementos"
         .Add "MainView.lbHorizontalSpace", "mm por ancho"
         .Add "MainView.lbVerticalSpace", "mm por altura"
        
         .Add "MainView.cbNewDocument", "Crear nuevo documento"
         .Add "MainView.cbGroup", "Composiciones de grupo"
        
         .Add "MainView.btnOk", "Inicio"

  End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
  Set ILocalizedStrings_Strings = Strings
End Property

