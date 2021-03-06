VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrRussian (1049)
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
    Set Strings = New Scripting.Dictionary
    With Strings
        
        .Add "Common.ErrNoDocument", "��� ��������� ���������"
        
        .Add "MainView.frForm", "������������ ����������"
        .Add "MainView.obByCount", "�� ����������"
        .Add "MainView.lbColumns", "��������"
        .Add "MainView.lbRows", "�����"
        .Add "MainView.opByArea", "�� �������"
        .Add "MainView.lbWidth", "�� �� ������"
        .Add "MainView.lbHeight", "�� �� ������"
        .Add "MainView.opByPage", "�� ������� ��������"
        
        .Add "MainView.frCompose", "������������ ����������"
        .Add "MainView.opOnePage", "�� ����� ��������"
        .Add "MainView.opMultiplePages", "������ �� ����� ��������"
        
        .Add "MainView.frSpaces", "���������� ����� ����������"
        .Add "MainView.lbHorizontalSpace", "�� �� ������"
        .Add "MainView.lbVerticalSpace", "�� �� ������"
        
        .Add "MainView.cbNewDocument", "������� ����� ��������"
        .Add "MainView.cbGroup", "������������ ����������"
        
        .Add "MainView.btnOk", "����"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
