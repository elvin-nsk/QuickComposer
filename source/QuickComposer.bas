Attribute VB_Name = "QuickComposer"
'===============================================================================
' ������                     : QuickComposer
' ������                     : 2022.02.21
' �����                        : https://vk.com/elvin_macro/QuickComposer
'                                        https://github.com/elvin-nsk
' �����                        : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "QuickComposer"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================

Private Const SpaceBetweenCompositionsDivider As Double = 10
Private Const CompositionsDocumentName As String = "Compositions"

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then
        VBA.MsgBox "��� ��������� ���������"
        Exit Sub
    End If
    
    Dim Cfg As Config
    Set Cfg = Config.CreateAndLoad
    ActiveDocument.Unit = cdrMillimeter
    
    If Not GetViewResult(Cfg) Then Exit Sub
    
    If Cfg.OptionNewDocument Then
        ActivePage.Shapes.All.CreateDocumentFrom.Activate
        ActiveDocument.Name = CompositionsDocumentName
    End If
    
    lib_elvin.BoostStart "QuickComposer", RELEASE
    Compose ActivePage.Shapes.All, ActivePage, ActiveDocument, Cfg
    ActiveDocument.Pages.First.Activate
    
Finally:
    lib_elvin.BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "������"
    Resume Finally

End Sub

'===============================================================================

Private Function GetViewResult(ByVal Cfg As Config) As Boolean
    With New MainView
    
        .obByCount = Cfg.OptionComposeByCount
        .opByArea = Cfg.OptionComposeByArea
        .opByPage = Cfg.OptionComposeByPage
        .opOnePage = Cfg.OptionOnePage
        .opMultiplePages = Cfg.OptionMultiplePages
        .cbNewDocument = Cfg.OptionNewDocument
        .cbGroup = Cfg.OptionGroup
        .Columns = Cfg.Columns
        .Rows = Cfg.Rows
        .Width = Cfg.Width
        .Height = Cfg.Height
        .HorizontalSpace = Cfg.HorizontalSpace
        .VerticalSpace = Cfg.VerticalSpace
        
        .Show
        GetViewResult = .IsOk
        If Not .IsOk Then Exit Function
        
        Cfg.OptionComposeByCount = .obByCount
        Cfg.OptionComposeByArea = .opByArea
        Cfg.OptionComposeByPage = .opByPage
        Cfg.OptionOnePage = .opOnePage
        Cfg.OptionMultiplePages = .opMultiplePages
        Cfg.OptionNewDocument = .cbNewDocument
        Cfg.OptionGroup = .cbGroup
        Cfg.Columns = .Columns
        Cfg.Rows = .Rows
        Cfg.Width = .Width
        Cfg.Height = .Height
        Cfg.HorizontalSpace = .HorizontalSpace
        Cfg.VerticalSpace = .VerticalSpace
    
    End With
End Function

Private Function Compose(ByVal Shapes As ShapeRange, _
                                                 ByVal Page As Page, _
                                                 ByVal Doc As Document, _
                                                 ByVal Cfg As Config)
    
    Dim MaxPlacesInWidth As Long
    Dim MaxPlacesInHeight As Long
    Dim MaxWidth As Double
    Dim MaxHeight As Double
    
    With Cfg
        If .OptionComposeByCount Then
            MaxPlacesInWidth = .Columns
            MaxPlacesInHeight = .Rows
            MaxWidth = 0
            MaxHeight = 0
        ElseIf .OptionComposeByArea Then
            MaxPlacesInWidth = 0
            MaxPlacesInHeight = 0
            MaxWidth = .Width
            MaxHeight = .Height
        ElseIf .OptionComposeByPage Then
            MaxPlacesInWidth = 0
            MaxPlacesInHeight = 0
            MaxWidth = Page.SizeWidth
            MaxHeight = Page.SizeHeight
        End If
    End With
    
    Dim Elements As Collection
    Set Elements = ShapesToElements(Shapes)
    Dim ComposedShapesAsElements As Collection
    Set ComposedShapesAsElements = New Collection
    
    Do
        With Composer.CreateAndCompose( _
                                        Elements:=Elements, _
                                        StartingPoint:=FreePoint.Create(-20000, 20000), _
                                        MaxPlacesInWidth:=MaxPlacesInWidth, _
                                        MaxPlacesInHeight:=MaxPlacesInHeight, _
                                        MaxWidth:=MaxWidth, _
                                        MaxHeight:=MaxHeight, _
                                        HorizontalSpace:=Cfg.HorizontalSpace, _
                                        VerticalSpace:=Cfg.VerticalSpace _
                                    )
            ComposedShapesAsElements.Add _
                ComposerElement.Create(ElementsToShapes(.ComposedElements))
            Set Elements = .RemainingElements
        End With
    Loop While Elements.Count > 0
    
    DistributeCompositions ComposedShapesAsElements, Doc, Cfg

End Function

Private Function DistributeCompositions _
                                    (ByVal Elements As Collection, _
                                     ByVal Doc As Document, _
                                     ByVal Cfg As Config)
    Dim i As Long
    Dim Count As Long
    Dim Space As Double
    Dim Shapes As ShapeRange
    If Cfg.OptionMultiplePages Then
        If Elements.Count > 1 Then Doc.AddPages Elements.Count - 1
        For i = 1 To Elements.Count
            If i > 1 Then _
                lib_elvin.MoveToLayer Elements(i).Shapes, Doc.Pages(i).ActiveLayer
            Elements(i).Shapes.CenterX = Doc.Pages(i).CenterX
            Elements(i).Shapes.CenterY = Doc.Pages(i).CenterY
            If Cfg.OptionGroup Then
                Doc.Pages(i).Activate
                Elements(i).Shapes.Group
            End If
        Next i
    ElseIf Cfg.OptionOnePage Then
        Count = VBA.Fix(VBA.Sqr(Elements.Count)) + 1
        Space = lib_elvin.AverageDim(Elements(1).Shapes) / _
                        SpaceBetweenCompositionsDivider
        With Composer.CreateAndCompose( _
                                        Elements:=Elements, _
                                        StartingPoint:=FreePoint.Create(-20000, 20000), _
                                        MaxPlacesInWidth:=Count, _
                                        MaxPlacesInHeight:=Count, _
                                        HorizontalSpace:=Space, _
                                        VerticalSpace:=Space _
                                    )
            With ElementsToShapes(.ComposedElements)
                .CenterX = Doc.ActivePage.CenterX
                .CenterY = Doc.ActivePage.CenterY
            End With
            If Cfg.OptionGroup Then GroupElementsShapes Elements
        End With
    End If
End Function

Private Sub GroupElementsShapes(ByVal Elements As Collection)
    Dim Element As ComposerElement
    For Each Element In Elements
        Element.Shapes.Group
    Next Element
End Sub

Private Function ShapesToElements(ByVal Shapes As ShapeRange) As Collection
    Dim Shape As Shape
    Set ShapesToElements = New Collection
    For Each Shape In Shapes
        ShapesToElements.Add ComposerElement.Create(Shape)
    Next Shape
End Function

Private Function ElementsToShapes _
                                 (ByVal ComposerElements As Collection) As ShapeRange
    Dim Item As ComposerElement
    Set ElementsToShapes = New ShapeRange
    For Each Item In ComposerElements
        ElementsToShapes.AddRange Item.Shapes
    Next Item
End Function

'===============================================================================
' �����
'===============================================================================

Private Sub testComposer()
    ActiveDocument.BeginCommandGroup "test"
    ActiveDocument.Unit = cdrMillimeter
    With Composer.CreateAndCompose( _
                                    Elements:=ShapesToElements(ActivePage.Shapes.All), _
                                    StartingPoint:=FreePoint.Create(0, 297), _
                                    MaxPlacesInWidth:=3, _
                                    MaxPlacesInHeight:=4, _
                                    MaxWidth:=0, _
                                    MaxHeight:=297, _
                                    HorizontalSpace:=0, _
                                    VerticalSpace:=0 _
                                )
        ElementsToShapes(.RemainingElements).ApplyNoFill
    End With
    ActiveDocument.EndCommandGroup
End Sub
