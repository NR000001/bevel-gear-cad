Attribute VB_Name = "ExportDrawing"
Option Explicit

' ===========================
' SolidWorks VBA 宏：
' 1) 打开 source/从动锥齿轮.SLDPRT
' 2) 创建 A3 横向工程图
' 3) 插入主视图、俯视图、右视图、等轴测视图
' 4) 添加参数说明文字
' 5) 导出 PDF、DXF、STEP
' ===========================

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swPart As SldWorks.ModelDoc2
Dim swDraw As SldWorks.DrawingDoc
Dim swExt As SldWorks.ModelDocExtension

Sub main()
    On Error GoTo EH

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "无法连接到 SolidWorks。", vbCritical
        Exit Sub
    End If

    Dim modelPath As String
    modelPath = GetPartPath()
    If modelPath = "" Then
        MsgBox "未找到模型文件：source/从动锥齿轮.SLDPRT", vbCritical
        Exit Sub
    End If

    ' 打开零件（只读方式，避免修改原始文件）
    Dim errs As Long, warns As Long
    Set swPart = swApp.OpenDoc6(modelPath, swDocumentTypes_e.swDocPART, _
                                swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", errs, warns)
    If swPart Is Nothing Then
        MsgBox "打开零件失败，错误码：" & CStr(errs), vbCritical
        Exit Sub
    End If

    Dim templatePath As String
    templatePath = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)
    If Len(Trim$(templatePath)) = 0 Then
        MsgBox "未配置默认工程图模板，请先在 SolidWorks 设置默认 Drawing 模板。", vbCritical
        Exit Sub
    End If

    ' 新建工程图
    Set swModel = swApp.NewDocument(templatePath, 0, 0, 0)
    If swModel Is Nothing Then
        MsgBox "创建工程图失败。", vbCritical
        Exit Sub
    End If
    Set swDraw = swModel

    ' A3 横向：宽 420mm，高 297mm（单位：米）
    Dim ok As Boolean
    ok = swDraw.SetupSheet5("Sheet1", swDwgPaperSizes_e.swDwgPapersUserDefined, _
                            swDwgTemplateNone, 1#, 1#, True, "", 0#, 0#, "", True, 0.42, 0.297, "Default")
    If Not ok Then
        MsgBox "设置 A3 横向图纸失败。", vbCritical
        Exit Sub
    End If

    swModel.ClearSelection2 True

    ' 插入主视图
    Dim vFront As SldWorks.View
    Set vFront = swDraw.CreateDrawViewFromModelView3(modelPath, "*Front", 0.14, 0.18, 0)
    If vFront Is Nothing Then
        MsgBox "插入主视图失败，请检查模型标准视图。", vbCritical
        Exit Sub
    End If

    ' 插入俯视图（与主视图投影视图关系）
    Dim vTop As SldWorks.View
    Set vTop = swDraw.CreateUnfoldedViewAt3(0.14, 0.27, 0, False)

    ' 插入右视图
    Dim vRight As SldWorks.View
    Set vRight = swDraw.CreateUnfoldedViewAt3(0.24, 0.18, 0, False)

    ' 插入等轴测视图
    Dim vIso As SldWorks.View
    Set vIso = swDraw.CreateDrawViewFromModelView3(modelPath, "*Isometric", 0.33, 0.24, 0)

    ' 添加参数说明文字（可按需要修改文字内容）
    swModel.ClearSelection2 True
    swModel.SetAddToDB True
    swModel.InsertNote "参数说明：" & vbCrLf & _
                       "1) 齿数 z = （请填写）" & vbCrLf & _
                       "2) 模数 m = （请填写）" & vbCrLf & _
                       "3) 压力角 α = （请填写）" & vbCrLf & _
                       "4) 材料 = （请填写）"
    swModel.SetAddToDB False

    Dim swNote As SldWorks.Note
    Set swNote = swModel.GetFirstAnnotation2
    If Not swNote Is Nothing Then
        swNote.SetTextPoint2 0.30, 0.05, 0
    End If

    ' 保存工程图
    Dim outDir As String
    outDir = GetOutputDir(modelPath)

    Dim drawPath As String
    drawPath = outDir & "从动锥齿轮工程图.SLDDRW"
    SaveDoc swModel, drawPath

    ' 导出 PDF / DXF
    SaveDoc swModel, outDir & "从动锥齿轮工程图.pdf"
    SaveDoc swModel, outDir & "从动锥齿轮工程图.dxf"

    ' 激活零件并导出 STEP
    swApp.ActivateDoc3 swPart.GetTitle, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, errs
    SaveDoc swPart, outDir & "从动锥齿轮.step"

    MsgBox "导出完成：" & vbCrLf & outDir, vbInformation
    Exit Sub

EH:
    MsgBox "宏运行异常：" & Err.Description, vbCritical
End Sub

Private Function GetPartPath() As String
    ' 优先使用宏目录上级/source/从动锥齿轮.SLDPRT
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim macroDir As String
    macroDir = fso.GetParentFolderName(swApp.GetCurrentMacroPathName)

    Dim p1 As String
    p1 = fso.BuildPath(fso.GetParentFolderName(macroDir), "source\从动锥齿轮.SLDPRT")
    If fso.FileExists(p1) Then
        GetPartPath = p1
        Exit Function
    End If

    ' 回退：仓库根目录/从动锥齿轮.SLDPRT
    Dim p2 As String
    p2 = fso.BuildPath(fso.GetParentFolderName(macroDir), "从动锥齿轮.SLDPRT")
    If fso.FileExists(p2) Then
        GetPartPath = p2
        Exit Function
    End If

    GetPartPath = ""
End Function

Private Function GetOutputDir(ByVal modelPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderPath As String
    folderPath = fso.GetParentFolderName(modelPath)

    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    GetOutputDir = folderPath
End Function

Private Sub SaveDoc(ByVal doc As SldWorks.ModelDoc2, ByVal targetPath As String)
    Dim errs As Long, warns As Long
    Dim ok As Boolean

    ok = doc.Extension.SaveAs(targetPath, swSaveAsVersion_e.swSaveAsCurrentVersion, _
                              swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, errs, warns)
    If Not ok Then
        MsgBox "保存失败：" & targetPath & vbCrLf & "错误码：" & CStr(errs), vbExclamation
    End If
End Sub
