' 页面初始化
Private Sub UserForm_Initialize()
    '始初始化字体列表 默认选中Arial字体
    cboFontFamily.AddItem "Arial"
    cboFontFamily.AddItem "黑体"
    cboFontFamily.AddItem "微软雅黑"
    cboFontFamily.ListIndex = 0

    ' 默认Excel数据填充到表格使用 先纵向后横向 的排列方式
    rdoRowFirst.Value = True
    rdoColFirst.Value = False

    ' 默认选中第一个功能页面
    MultiPage1.Value = 0

    ' 默认Excel填充不可用
    btnDoExcelFill.Enabled = False
End Sub


' 每次页面激活时，如果只选中一个对象，则读取该物体的宽度高度作为批量设置尺寸的宽高值
Private Sub UserForm_Activate()
    ' On Error Resume Next
    If Not ActiveDocument Is Nothing Then
        ActiveDocument.Unit = cdrMillimeter

        grpSameSize.Caption = "批量设置尺寸 当前选中数量: " & ActiveSelection.Shapes.Count
        grpSameRotate.Caption = "批量设置旋转 当前选中数量: " & ActiveSelection.Shapes.Count
            
        If ActiveSelection.Shapes.Count > 0 Then
            btnSameRotateApply.Enabled = True
        
            Dim initWidth As Double, initHeight As Double
            If ActiveSelection.Shapes.Count = 1 Then
                txtWidth.Value = ActiveSelection.Shapes(1).SizeWidth
                txtHeight.Value = ActiveSelection.Shapes(1).SizeHeight

                ' 根据页面大小自动计算 使用Excel数据填充到表格功能下 页面上最多能放置的行数和列数
                Dim TableShape As Shape
                Set TableShape = ActiveSelection.Shapes(1)
                If TableShape.Type = cdrCustomShape Then
                    txtRowCount.Value = Int(txtPageHeight.Value / (TableShape.SizeHeight + txtRowSpan.Value))
                    txtColCount.Value = Int(txtPageWidth.Value / (TableShape.SizeWidth + txtColSpan.Value))

                    ' 启用Excel填充
                    btnDoExcelFill.Enabled = True
                Else
                    ' 禁用Excel填充
                    btnDoExcelFill.Enabled = False
                End If
            End If
            If ActiveSelection.Shapes.Count = 0 Then
                txtWidth.Value = 0
                txtHeight.Value = 0

                ' 禁用Excel填充
                btnDoExcelFill.Enabled = False
            End If
        Else
            btnSameRotateApply.Enabled = False
        End If
        
    End If
End Sub

' 点击按钮关闭插件
Private Sub btnCancel_Click()
    'Unload Me
    ' 隐藏而不关闭
    Me.Hide
End Sub

' 创建新雕刻幅面 界面操作
Private Sub btnCreateDoc_Click()
    Dim doc1 As Document
    Set doc1 = CreateDocument
    Call SetPageSettings
End Sub

' 设置雕刻幅面 界面操作
Private Sub btnSetDoc_Click()
    On Error Resume Next
    Call SetPageSettings
End Sub

' 设置雕刻幅面 功能函数
Private Sub SetPageSettings()
    Dim pageWidth As Double, pageHeight As Double
    pageWidth = txtPageWidth.Value
    pageHeight = txtPageHeight.Value

    ' 设置页面单位为毫米
    ActiveDocument.Unit = cdrMillimeter
    With ActiveDocument.MasterPage
        .SetSize pageWidth, pageHeight
        .Orientation = cdrLandscape
        .PrintExportBackground = True
        .Bleed = 0#
        .Background = cdrPageBackgroundNone
    End With
End Sub

' 生成雕刻表格 界面操作
Private Sub btnCreateCarveTable_Click()
    'https://community.coreldraw.com/sdk/api/draw/18/m/layer.createcustomshape
    'https://community.coreldraw.com/sdk/f/code-snippets-feedback/51740/tables-all-sorts-of-slow
    
    Call CreateTable
    
    'Call ChangeClipData
End Sub

' 生成雕刻表格 功能实现
Sub CreateTable()
    ' 定义命令组 用于整体撤销
    ActiveDocument.BeginCommandGroup "CarvePlugin - 生成雕刻表格"

        ' 对应手动操作的步骤为：编辑->选择性粘贴->Rich Text Format->摒弃字体和格式->将表格导入为表格
        ActiveLayer.PasteSpecial "Rich Text Format"
        
        On Error GoTo ErrHandler
    
        Dim PastedShapeRange As ShapeRange
        Set PastedShapeRange = ActiveSelectionRange
            
        Dim ExcelTableShape As Shape
        ' 粘贴后第一层级为对象群组 第二层级为表格
        Set ExcelTableShape = PastedShapeRange.Shapes(1).Shapes(1)
        
        Dim RowCount As Integer, ColumnCount As Integer
        RowCount = ExcelTableShape.Custom.Rows.Count
        ColumnCount = ExcelTableShape.Custom.Columns.Count
            
        Dim cellWidth As Double, cellHeight As Double
        Dim tableWidth As Double, tableHeight As Double
        cellWidth = txtCellWidth.Value
        cellHeight = txtCellHeight.Value
        tableWidth = cellWidth * ColumnCount
        tableHeight = cellHeight * RowCount
        
        ActiveDocument.ReferencePoint = cdrCenter
        ActiveDocument.Unit = cdrMillimeter
        
        Set createdTable = ActiveLayer.CreateCustomShape("Table", 10, 10, tableWidth + 10, tableHeight + 10, ColumnCount, RowCount)
    
        Optimization = True
        For i = 1 To RowCount
            For j = 1 To ColumnCount
                textcontent = ExcelTableShape.Custom.Cell(j, i).TextShape.Text.Story
                If textcontent <> "" Then
                    ' 设置每个单元格的和表格之间的间隙都为0 避免单元格较小时默认2mm的空隙会导致文本无法录入
                    Dim tableCell As Variant
                    Set tabelCell = createdTable.Custom.Cell(j, i)
                    tabelCell.SetAllMargins 0

                    ' 文本内容
                    tabelCell.TextShape.Text.Story = textcontent
                    ' 字体类型
                    tabelCell.TextShape.Text.Story.Font = cboFontFamily.Text
                    ' 字体大小
                    tabelCell.TextShape.Text.Story.Words.All.Size = txtCellFontSize.Value
                    ' 默认单元格内水平居中
                    tabelCell.TextShape.Text.Story.Alignment = cdrCenterAlignment
                    ' 默认单元格内垂直居中 界面操作在 菜单->文本->段落格式化(P)
                    tabelCell.TextShape.Text.Frame.VerticalAlignment = cdrCenterJustify
                End If
                 
            Next j
        Next i

        '所有框线红色 红色对应激光雕刻驱动中的切割
        createdTable.Custom.borders.All.Color = CreateCMYKColor(0, 100, 100, 0)
        '所有框线毛细
        createdTable.Custom.borders.All.width = 0.0762
            
        ' 删除最初粘贴的部分
        ExcelTableShape.Delete
        
        ' 居中放置新建的 CustomShape(Table)
        createdTable.AlignToPageCenter cdrAlignLeft + cdrAlignRight + cdrAlignTop + cdrAlignBottom, cdrTextAlignBoundingBox
        
        ' 设置表格整体大小
        createdTable.SizeWidth = tableWidth
        createdTable.SizeHeight = tableHeight

    ' 结束命令组
    ActiveDocument.EndCommandGroup    
    Me.Hide
    
ExitSub:
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub
ErrHandler:
    MsgBox "出现错误: " & Err.Description
    Resume ExitSub
End Sub


' 统一行间距 该功能其实无用 通过在菜单->文本->段落格式化(P)中也可以实现
Private Sub btnSameLineSpacing_Click()
    On Error Resume Next
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter

    If ActiveSelection.Shapes.Count > 0 Then
        On Error GoTo ErrHandler

        Dim LineSpacing As Single, ParaBefore As Single, ParaAfter As Single
        LineSpacing = txtLineSpacing.Value
        ParaBefore = txtParaBefore.Value
        ParaAfter = txtParaAfter.Value

        '开始命令组
        ActiveDocument.BeginCommandGroup "CarvePlugin - 统一行间距"
        
        Set TableShape = ActiveSelection.Shapes(1)  ' 表格(TableShape) cdrCustomShape -> 单元格 cdrCustomShape -> 段落文本 cdrTextShape
        If TableShape.Type = cdrCustomShape Then
            Dim RowCount As Integer, ColumnCount As Integer
            RowCount = TableShape.Custom.Rows.Count
            ColumnCount = TableShape.Custom.Columns.Count
            
            Optimization = True
            For i = 1 To RowCount
                For j = 1 To ColumnCount
                    Dim s As Shape
                    Set s = TableShape.Custom.Cell(j, i).TextShape ' 整个单元格的文本框
                    If s.Type = cdrTextShape Then
                        Dim t As Text
                        Set t = s.Text
                        If t.Type = cdrParagraphText Then
                            ' t.Story.ParaSpacingAfter
                            '在界面上显示默认为 前100 后0 行100 （传入参数顺序为 行100 前100 后0）
                            t.Story.SetLineSpacing cdrPercentOfCharacterHeightLineSpacing, LineSpacing, ParaBefore, ParaAfter
                            
                            ' t.Story.Paragraphs.Item(1).Bold = True
                            ' t.Story.Paragraphs.Item(1).Underline = cdrDoubleThinFontLine
                            ' Set AllParagraphs = t.Story.Paragraphs
                            ' For k = 1 To AllParagraphs.Count
                            '     Set oneParagraph = AllParagraphs.Item(k)
                            '     oneParagraph.Underline = cdrDoubleThinFontLine

                            '     Set tmp = oneParagraph.Characters.Item(1)  ' 表示其中的每个字符
                            '     tmp.Underline = cdrDoubleThinFontLine

                            '     with oneParagraph
                            '         .Bold = True
                            '         .Italic = True
                            '         .Underline = cdrDoubleThinFontLine
                            '     End With
                            ' Next k

                        End If
                    End If
                Next j
            Next i
        End If
        
        '结束命令组
        ActiveDocument.EndCommandGroup        
        Me.Hide

    Else
        MsgBox "未选择修改对象!", vbOKOnly + vbCritical, "错误"
    End If
ExitSub:
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub
ErrHandler:
    MsgBox "出现错误: " & Err.Description
    Resume ExitSub
End Sub


' 自动文字尺寸
Private Sub btnAutoTextSizeApply_Click()

    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    '开始命令组
    ActiveDocument.BeginCommandGroup "CarvePlugin - 自动文字尺寸"
    
    Dim brk As ShapeRange
    Set brk = ActiveSelection.BreakApartEx
    '转曲线
    brk.ConvertToCurves
    
    Dim grp As ShapeRange
    '解散群组
    Set grp = brk.UngroupEx
    
    
    Dim textWidthPercent As Double, textHeigthPercent As Double
    textWidthPercent = txtTextPercent.Value / 100#
    textHeigthPercent = txtTextPercent.Value / 100#
    
    Dim cellWidth As Double, cellHeight As Double
    cellWidth = txtCellWidth.Value
    cellHeight = txtCellHeight.Value
    
    Dim allShape As ShapeRange, eachShape As Shape
    Set allShape = grp
    
    Optimization = True
    If allShape.Count > 0 Then
        For k = 1 To allShape.Count
        
            Set tempShape = allShape.Item(k)
            
            Dim oldTextWidth As Double, oldTextHeight As Double
            oldTextWidth = tempShape.SizeWidth
            oldTextHeight = tempShape.SizeHeight
            
            Dim initTextWidth As Double, initTextHeight As Double
            initTextWidth = cellWidth * textWidthPercent
            initTextHeight = cellHeight * textHeigthPercent
            
            Dim initTextWidthScalePercent As Double, initTextHeightScalePercent As Double
            initTextWidthScalePercent = initTextWidth / oldTextWidth
            initTextHeightScalePercent = initTextHeight / oldTextHeight
            
            Dim LastScalePercent As Double
            If initTextWidthScalePercent <= initTextHeightScalePercent Then
                LastScalePercent = initTextWidthScalePercent
            Else
                LastScalePercent = initTextHeightScalePercent
            End If
        
            'Curve
            If tempShape.Type = cdrCurveShape Then
                If tempShape.SizeWidth > 0.001 And tempShape.SizeHeight > 0.001 Then  'Is Text Curve
                    tempShape.SizeWidth = oldTextWidth * LastScalePercent
                    tempShape.SizeHeight = oldTextHeight * LastScalePercent
                End If
                'Debug.Print "Curve: Width: " & tempShape.SizeWidth & " Height: " & tempShape.SizeHeight
            End If
            

            'Group
            If tempShape.Type = cdrGroupShape Then
                tempShape.SizeWidth = oldTextWidth * LastScalePercent
                tempShape.SizeHeight = oldTextHeight * LastScalePercent
                'Debug.Print "Group"
            End If
    
        Next k
    End If
    
    '结束命令组
    ActiveDocument.EndCommandGroup    
    Me.Hide
                
ExitSub:
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub
ErrHandler:
    MsgBox "出现错误: " & Err.Description
    Resume ExitSub
End Sub



'批量设置旋转
Private Sub btnSameRotateApply_Click()
    On Error Resume Next
    
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    Dim RotateAngle As Double
    RotateAngle = txtRotateAngle.Value
    
    If ActiveSelection.Shapes.Count > 0 Then
        On Error GoTo ErrHandler
        Optimization = True
    
        '开始命令组
        ActiveDocument.BeginCommandGroup "CarvePlugin - 批量设置旋转"
            
            Dim brk As ShapeRange
            Set brk = ActiveSelection.BreakApartEx
            If brk.Count > 0 Then
                For k = 1 To brk.Count
                
                    Set tempShape = brk.Item(k)
                    
                    'Curve
                    If tempShape.Type = cdrCurveShape Then
                        If tempShape.SizeWidth > 0.001 And tempShape.SizeHeight > 0.001 Then  'Is Text Curve
                            tempShape.Rotate RotateAngle
                        End If
                    End If
    
                    'Group
                    If tempShape.Type = cdrGroupShape Then
                        tempShape.Rotate RotateAngle
                    End If
                    
                    
                    'cdrTextShape
                    If tempShape.Type = cdrTextShape Then
                        tempShape.Rotate RotateAngle
                    End If
                    
                Next k
            End If
        
        '结束命令组
        ActiveDocument.EndCommandGroup        
        Me.Hide
    Else
        MsgBox "未选择旋转对象!", vbOKOnly + vbCritical, "错误"
    End If
    
ExitSub:
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub
ErrHandler:
    MsgBox "出现错误: " & Err.Description
    Resume ExitSub
End Sub


'批量设置尺寸
Private Sub btnSameSizeApply_Click()
    On Error Resume Next
    
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    
    Dim width As Double, height As Double
    width = txtWidth.Value
    height = txtHeight.Value
    
    If ActiveSelection.Shapes.Count > 0 Then
        If width <> 0 And height <> 0 Then
            On Error GoTo ErrHandler
            Optimization = True
        
            '开始命令组
            ActiveDocument.BeginCommandGroup "CarvePlugin - 批量设置尺寸"
                
                Dim brk As ShapeRange
                Set brk = ActiveSelection.BreakApartEx
                If brk.Count > 0 Then
                    For k = 1 To brk.Count
                    
                        Set tempShape = brk.Item(k)
                        
                        'Curve
                        If tempShape.Type = cdrCurveShape Then
                            If tempShape.SizeWidth > 0.001 And tempShape.SizeHeight > 0.001 Then  'Is Text Curve
                                tempShape.SizeWidth = width
                                tempShape.SizeHeight = height
                            End If
                            'Debug.Print "Curve: Width: " & tempShape.SizeWidth & " Height: " & tempShape.SizeHeight
                        End If
        
                        'Group
                        If tempShape.Type = cdrGroupShape Then
                            tempShape.SizeWidth = width
                            tempShape.SizeHeight = height
                            'Debug.Print "Group"
                        End If
                        
                    Next k
                End If
            
            ActiveDocument.EndCommandGroup
            '结束命令组            
            'MsgBox "共修改 " + CStr(ActiveSelection.Shapes.Count) + " 个对象!", vbOKOnly + vbInformation, "完成"
            Me.Hide

        Else
            If width = 0 Then
                MsgBox "宽度不能设置为0!", vbOKOnly + vbCritical, "错误"
            Else
                MsgBox "高度不能设置为0!", vbOKOnly + vbCritical, "错误"
            End If
        
        End If
    Else
        MsgBox "未选择修改对象!", vbOKOnly + vbCritical, "错误"
    End If
    
ExitSub:
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub
ErrHandler:
    MsgBox "出现错误: " & Err.Description & " 请检查Excel数据!"
    Resume ExitSub
End Sub


' Excel数据填充到表格
Private Sub btnDoExcelFill_Click()
    On Error Resume Next
    ' 设置文档单位
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter


    ' 检查是否选择了对象
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "未选择任何操作对象!", vbOKOnly + vbCritical, "错误"
        Exit Sub
    End If

    ' 检查是否选择了对象
    If ActiveSelection.Shapes.Count > 1 Then
        MsgBox "只能选择一个表格对象!", vbOKOnly + vbCritical, "错误"
        Exit Sub
    End If

    Dim srcShape As Shape
    ' 获取当前选中的对象(仅第一个对象)
    Set srcShape = ActiveSelection.Shapes(1)


    ' 选择Excel文件
    Dim excelApp As Object
    Set excelApp = CreateObject("Excel.Application")

    ' ' 创建一个 Shell.Application 对象
    ' Dim shell As Object
    ' Set shell = CreateObject("Shell.Application")

    ' ' 获取桌面路径
    ' Dim desktopPath As String
    ' desktopPath = shell.Namespace(10).Self.Path
    ' MsgBox desktopPath

    ' Dim specialFolderPath As String
    ' specialFolderPath = shell.Namespace(desktopPath).Self.Path
    ' MsgBox specialFolderPath

    ' 选择Excel文件
    Dim ExcelFileToOpen As String
    ' 设置当前文件夹（有可能不生效）
    SETCURRFOLDER = "C:\"
    ' 使用CorelDraw原生打开文件对话框
    ' 参看 https://www.cdrvba.com/article-coreldraw-vba-open-file-dialog
    ExcelFileToOpen = CorelScriptTools.GetFileBox("所有 Excel 文件(*.xls*)|*.xls;*.xlsx|所有文件(*.*)|*.*", "请选择Excel数据文件", 0, "")
    ' 使用Excel功能里的打开文件对话框 该对话框非模态显示 可能会出现在后台 被遮挡
    'ExcelFileToOpen = excelApp.GetOpenFilename("所有 Excel 文件(*.xls*),*.xls;*.xlsx,所有文件(*.*),*.*")
    
    If ExcelFileToOpen = "" Then
        Set excelApp = Nothing
        MsgBox "未选择Excel数据文件!", vbOKOnly + vbCritical, "错误"
        Exit Sub
    End If


    Dim bRowFirst As Boolean, bColFirst As Boolean
    Dim RowCount As Long, ColCount As Long
    Dim RowSpan As Double, ColSpan As Double

    bRowFirst = rdoRowFirst.Value
    bColFirst = rdoColFirst.Value
    RowCount = txtRowCount.Value
    ColCount = txtColCount.Value
    RowSpan = txtRowSpan.Value
    ColSpan = txtColSpan.Value

    Dim TotalCounter As Long

    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim cellValue As Variant

    If ActiveSelection.Shapes.Count = 1 Then
        On Error GoTo ErrHandler

        ' 打开 Excel 工作簿
        Set excelWorkbook = excelApp.Workbooks.Open(ExcelFileToOpen)
        
        ' 指定要操作的工作表 只处理第一个工作薄
        Set excelWorksheet = excelWorkbook.Worksheets(1)


        '开始命令组 模板填充
        ActiveDocument.BeginCommandGroup "CarvePlugin - Excel数据填充到表格"

        Optimization = True
        TotalCounter = 1
        If bRowFirst Then
            For j = 1 To ColCount
                For i = 1 To RowCount
                    Set TableShape = DuplicateShape(srcShape, i, j, ColSpan, RowSpan)
                    Call ReplaceShapeText(TableShape, TotalCounter, excelWorksheet)
                    TotalCounter = TotalCounter + 1
                Next i
            Next j
        End If

        If bColFirst Then
            For i = 1 To RowCount
                For j = 1 To ColCount
                    Set TableShape = DuplicateShape(srcShape, i, j, ColSpan, RowSpan)
                    Call ReplaceShapeText(TableShape, TotalCounter, excelWorksheet)
                    TotalCounter = TotalCounter + 1
                Next j
            Next i
        End If

        '结束命令组
        ActiveDocument.EndCommandGroup

        ' 保存 Excel 工作簿
        excelWorkbook.Save
        ' 关闭 Excel 工作簿
        excelWorkbook.Close
        ' 退出 Excel 应用程序
        excelApp.Quit
        
        ' 释放对象
        Set excelWorksheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
    End If

    Me.Hide

ExitSub:
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub
ErrHandler:
    MsgBox "出现错误: " & Err.Description
    Resume ExitSub
End Sub


' 替换Shape中的文本为Excel单元格数据
Private Function ReplaceShapeText(ByVal TableShape As Shape, ByVal TotalCounter As Long, ByVal excelWorksheet As Object)
    If TableShape.Type = cdrCustomShape Then
        ' 获取表格的行数
        Dim RowCount As Integer
        RowCount = TableShape.Custom.Rows.Count ' 用于计算在Excel中应当对应的行数

        ' 创建正则表达式对象
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Global = True ' 全局匹配
            .Pattern = "\{([A-Z]{1,3})(\d+)\}" '设置模式 用于匹配类似于 {A1} {B2} {AA21} 这样的文本
        End With


        Dim element As Variant
        Dim fsize As Double
        Dim textcontent As String
        For Each element In TableShape.Custom.Cells ' 遍历表格中的每个单元格
            'If Not element Is Nothing Then
            Dim t As Text
            Set t = element.TextShape.Text
            ' 获取单元格内的文本内容及字体大小
            textcontent = t.Story.Text
            fsize = t.Story.Size

            Dim matches As Object
            Set matches = regex.Execute(textcontent)

            Dim match As Object
            For Each match In matches
                If match.SubMatches.Count = 2 Then
                    Dim letterPart As String
                    letterPart = match.SubMatches(0)
                    
                    Dim numberPart As String
                    numberPart = match.SubMatches(1)

                    Dim fomuer_number As Integer
                    fomuer_number = CInt(numberPart)
                    Dim real_number As Integer
                    real_number = RowCount * (TotalCounter - 1) + fomuer_number

                    ' 取Excel单元格的值 （如果使用Value，则无法正确读取使用 设置单元格格式 后表示的数据，只会读取到原始的值）
                    cellValue = excelWorksheet.Range(letterPart & CStr(real_number)).Text

                    ' Excel中的单元格文本如包含换行符，则获取到的文本仅含有Chr(10)字符而不是vbCrLf或者vbNewLine，因此需要特殊处理
                    Dim arr() As String
                    Dim cellValueWithNewLine As String
                    cellValueWithNewLine = ""
                    arr = Split(cellValue, Chr(10))
                    If UBound(arr) > 0 Then
                        For i = 0 To UBound(arr)
                            If i <> UBound(arr) Then
                                cellValueWithNewLine = cellValueWithNewLine & arr(i) & vbCrLf
                            Else
                                cellValueWithNewLine = cellValueWithNewLine & arr(i)
                            End If
                        Next i
                    Else
                        cellValueWithNewLine = cellValue
                    End If

                    ' 替换原文本的内容
                    textcontent = Replace(textcontent, "{" & letterPart & numberPart & "}", cellValueWithNewLine)

                    ' 设置已读取数据的Excel单元格的背景颜色为红色
                    excelWorksheet.Range(letterPart & CStr(real_number)).Interior.Color = RGB(255, 0, 0)
                End If
            Next match
            ' 到这里已经替换结束 更新文本内容
            t.Story.Text = textcontent
            ' 再次把之前获取到的字体大小设置到元素中，不这样处理的话中文字体大小会不统一
            t.Story.Words.All.Size = fsize

        Next element

    End If
End Function


' 复制对象到行列位置 参数：对象，行号，列号，行间距(单位毫米)，列间距(单位毫米), 需要在调用的上层过程中设置文档单位为毫米
Private Function DuplicateShape(ByVal srcShape As Shape, ByVal RowNumber As Integer, ByVal ColumnNumber As Integer, ByVal RowSpan As Double, ByVal ColSpan As Double) As Shape
    Dim duplicatedShape As Shape
    Dim SizeHeight As Double, SizeWidth As Double
    SizeHeight = srcShape.SizeHeight
    SizeWidth = srcShape.SizeWidth

    Set ds = srcShape.Duplicate
    ds.Move (SizeWidth + ColSpan) * (ColumnNumber - 1), -(SizeHeight + RowSpan) * RowNumber
    Set DuplicateShape = ds
End Function


' 缩放活动视图到匹配页面
Private Sub ZoomToFitPage()
    Dim pageWidth As Double, pageHeight As Double
    pageWidth = txtPageWidth.Value
    pageHeight = txtPageHeight.Value
    ActiveWindow.ActiveView.ToFitArea -5, -5, pageWidth + 5, pageHeight + 5
End Sub



Sub ChangeClipData()
    '开始命令组
    ActiveDocument.BeginCommandGroup "CarvePlugin - 生成雕刻表格"
            
        ActiveLayer.PasteSpecial "Rich Text Format"
        
        On Error GoTo ErrHandler
        
        Dim PastedShapeRange As ShapeRange
        Set PastedShapeRange = ActiveSelectionRange
        
        'If PastedShapeRange.Shapes(1).Type = cdrShapeType.cdrGroupShape Then
             'Debug.Print PastedShapeRange.Shapes(1).Type
        'End If
        
        ActiveDocument.ReferencePoint = cdrCenter
        ActiveDocument.Unit = cdrMillimeter
        
        Dim ExcelTableShape As Shape
        Set ExcelTableShape = PastedShapeRange.Shapes(1).Shapes(1)
        
        Dim RowCount As Integer, ColumnCount As Integer
        RowCount = ExcelTableShape.Custom.Rows.Count
        ColumnCount = ExcelTableShape.Custom.Columns.Count
        
        Dim cellWidth As Double, cellHeight As Double
        cellWidth = txtCellWidth.Value
        cellHeight = txtCellHeight.Value
        
        '宽高需要多执行几次保证其数据
        PastedShapeRange.SizeWidth = cellWidth * ColumnCount
        PastedShapeRange.SizeHeight = cellHeight * RowCount
        
        '页面居中
        PastedShapeRange.AlignToPageCenter cdrAlignLeft + cdrAlignRight + cdrAlignTop + cdrAlignBottom, cdrTextAlignBoundingBox
        
        '宽高需要多执行几次保证其数据
        PastedShapeRange.SizeWidth = cellWidth * ColumnCount
        PastedShapeRange.SizeHeight = cellHeight * RowCount
        
        
        On Error GoTo ErrHandler
        Optimization = True
        
        With ExcelTableShape.Custom
            For i = 1 To .Rows.Count
                For j = 1 To .Columns.Count
                    '.Cell(j, i).TextShape.Text.Story = "Sun"
                    textcontent = .Cell(j, i).TextShape.Text.Story
                    'MsgBox textcontent
                    
                    If .Cell(j, i).TextShape.Text.Story <> "" Then
                       .Cell(j, i).TextShape.Text.Story.Font = "Arial"
                       .Cell(j, i).TextShape.Text.Story.Words.All.Size = txtCellFontSize.Value
                       .Cell(j, i).TextShape.Text.Story.Alignment = cdrCenterAlignment
                       .Cell(j, i).TextShape.Text.Frame.VerticalAlignment = cdrCenterJustify
                    End If
                     
                Next j
            Next i
        End With
        
        '所有框线红色
        ExcelTableShape.Custom.borders.All.Color = CreateCMYKColor(0, 100, 100, 0)
        '所有框线毛细
        ExcelTableShape.Custom.borders.All.width = 0.0762
        
        
        '宽高需要多执行几次保证其数据
        PastedShapeRange.SizeWidth = cellWidth * ColumnCount
        PastedShapeRange.SizeHeight = cellHeight * RowCount
    
    ActiveDocument.EndCommandGroup
    Me.Hide
    
ExitSub:
    Optimization = False
    ActiveWindow.Refresh
    Exit Sub
ErrHandler:
    MsgBox "出现错误: " & Err.Description
    Resume ExitSub
End Sub