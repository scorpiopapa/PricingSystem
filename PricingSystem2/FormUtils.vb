Imports Microsoft.Office.Interop
Imports System.Data.OleDb

Module FormUtils
#If DEBUG Then
    'Public TEMPLATE_FILE As String = "..\..\..\报价模板.xls"
    'Public REPORT_FILE As String = "..\..\..\报价单位A 项目1 2014年报价.xls"
    'Public REPORT_FILE_BAK As String = "..\..\..\报价单位A 项目1 2014年报价"
    Public BLANK_FILE As String = Application.StartupPath + "\..\..\..\blank.xls"
    'Public BLANK_FILE As String = "F:\Project\供应商系统\PricingSystem2\PricingSystem2\blank.xls"
#Else
    'Public TEMPLATE_FILE As String = Application.StartupPath + "\config\报价模板.xls"
    'Public REPORT_FILE As String = Application.StartupPath + "\config\报价单位A 项目1 2014年报价.xls"
    'Public REPORT_FILE_BAK As String = Application.StartupPath + "\config\报价单位A 项目1 2014年报价"
    'Public TEMPLATE_FILE As String = Application.StartupPath + "\config\temp.xls"
    Public BLANK_FILE As String = Application.StartupPath + "\config\blank.xls"
#End If

    Public ReadOnly ERR_TITLE As String = "错误"
    Public ReadOnly CONNECT_DB_ERR As String = "读取数据库失败！"

    'Public Shared ReadOnly EXPORT_NO_SELECT As String = "请至少选择一条记录"
    'Public Shared ReadOnly DELETE_NO_SELECT As String = "请至少选择一条记录"
    'Public Shared ReadOnly UPDATE_NO_SELECT As String = "请至少选择一条记录"
    'Public Shared ReadOnly EXPORT_WITH_UNREVIEWED As String = "不能导出未审价记录"

    'Public ReadOnly QUERY_DATE_ERR As String = "开始日期不能大于结束日期"

    Public ReadOnly INFO_TITLE As String = "提示"
    Public ReadOnly QUESTION_TITLE As String = "确认"
    'Public ReadOnly EXPORT_DONE As String = "{0}条记录导出完成，是否打开导出目录？"
    'Public ReadOnly DELETE_ROWS As String = "是否要删除选中的{0}条记录？"
    'Public ReadOnly IMPORT_CONFIRM As String = "导入的记录将会覆盖本地原有数据，是否要导入？"

    'Public Shared ReadOnly DELETE_DONE As String = "{0}条记录删除成功"
    'Public Shared ReadOnly NO_RECORD As String = "没有检索到记录"
    'Public Shared ReadOnly IMPORT_DONE As String = "{0}条记录导入成功"
    'Public Shared ReadOnly UPDATE_DONE As String = "当前记录修改成功"

    'Public ReadOnly EXCEL_NO_MATCH As String = "报表类型不匹配"

    Public ReadOnly EMPTY_COM_NAME As String = "单位名称不能为空"
    Public ReadOnly EMPTY_PROJECT_NAME As String = "项目名称不能为空"
    Public ReadOnly EMPTY_EXPORT_ITEM As String = "请选择要导出的项目名称或报表名称"

    Public ReadOnly EMPTY_REPORT_NAME As String = "报表名称不能为空"
    Public ReadOnly REPORT_EXIST As String = "报表已存在"
    Public ReadOnly REPORT_UPDATE_DONE As String = "报表更新成功"
    Public ReadOnly REPORT_ADD_DONE As String = "报表添加成功"
    Public ReadOnly EMPTY_REPORT_LIST As String = "请添加要导出的报表"
    Public ReadOnly REPORT_DATA_UPDATE_DONE As String = "数据更新成功"
    Public ReadOnly REPORT_DATA_EXPORT_DONE As String = "报表导出成功"
    Public ReadOnly REPORT_DATA_IMPORT_DONE As String = "报表导入成功"

    Public ReadOnly REPORT_NAME_ROW As Integer = 2
    Public ReadOnly REPORT_NAME_FONT_SIZE As Integer = 20
    Public ReadOnly MARKER_ROW As Integer = 4
    Public ReadOnly HEAD_ROW As Integer = MARKER_ROW + 1
    Public ReadOnly SUMMARY_ROW As Integer = HEAD_ROW + 1
    Public ReadOnly DATA_START_ROW As Integer = SUMMARY_ROW + 1
    Public ReadOnly EXCEL_ITEM_ID_COLUMN As Integer = 1
    Public ReadOnly START_COLUMN As Integer = 2
    Public ReadOnly NO_TEXT As String = "序号"
    Public ReadOnly SUMMARY_TEXT As String = "总计"
    Public ReadOnly MARKER_TEXT1 As String = "项目名称："
    Public ReadOnly MARKER_TEXT2 As String = "金额单位：元"
    Public ReadOnly BID_PRICE_TEXT As String = "报价"
    Public ReadOnly OFFER_PRICE_TEXT As String = "审价"

    Public ReadOnly YEAR_TEXT As String = "年度"
    Public ReadOnly REPORT_SUMMARY_TEXT As String = "报表总览"

    Public Sub ShowErrorMessage(msg As String)
        MsgBox(msg, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, ERR_TITLE)
    End Sub

    Public Sub ShowInfoMessage(msg As String)
        MsgBox(msg, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, INFO_TITLE)
    End Sub

    Public Function ShowQuestionMessage(msg As String) As MsgBoxResult
        Return MsgBox(msg, MsgBoxStyle.Question + MsgBoxStyle.YesNo, QUESTION_TITLE)
    End Function

    Public Function Empty(s As String) As Boolean
        Return IsNothing(s) OrElse s.Length = 0
    End Function

    Public Function FormatDate(d As Date) As String
        'Return Format(d, "yyyy年MM月dd日 HH时mm分ss秒")
        Return Format(d, "yyyy-MM-dd")
    End Function

    Public Sub ShowMainForm(frm As Form)
        With MainForm
            If TypeOf frm Is BuyerLogin Then
                .LoginUserName = CType(frm, BuyerLogin).TextBox1.Text
            Else
                .LoginUserName = CType(frm, SellerLogin).TextBox1.Text
                .ComName = CType(frm, SellerLogin).TextBox3.Text
            End If

            .LoginForm = frm
            frm.Hide()
            .Show()
        End With
    End Sub

    Public Function ExcelFilter() As String
        Return String.Format("Excel 2003 (*{0})|*{0}", ".xls")
    End Function

    Public Sub BindComboBoxData(cb As ComboBox, table As DataTable, bindColumnName As String, Optional style As ComboBoxStyle = ComboBoxStyle.DropDown)
        With cb
            .DropDownStyle = style
            .AutoCompleteMode = AutoCompleteMode.Suggest
            .AutoCompleteSource = AutoCompleteSource.ListItems

            '.Text = ""
            .DisplayMember = bindColumnName
            '.ValueMember = bindColumnName
            .DataSource = table
            .Text = ""
        End With
    End Sub

    Public Sub AddSumFormula(r As Excel.Range)
        r.FormulaR1C1 = "=SUM(R[1]C:R[60000]C)"
    End Sub

    Public Function Group(data As List(Of List(Of String))) As Dictionary(Of String, List(Of List(Of String)))
        Dim dic As New Dictionary(Of String, List(Of List(Of String)))

        For Each l As List(Of String) In data
            Dim key As String = l(0)
            Dim values As List(Of List(Of String))

            If dic.ContainsKey(key) Then
                values = dic.Item(key)
            Else
                values = New List(Of List(Of String))
                dic.Add(key, values)
            End If

            Dim tmp As New List(Of String)
            tmp.AddRange(l)
            tmp.RemoveAt(0)

            values.Add(tmp)
        Next

        Return dic
    End Function

    Public Function FindReportNames(node As TreeNode) As List(Of String)
        Dim names As New List(Of String)

        If node.Nodes.Count = 0 Then
            names.Add(node.Text)
        Else
            'Dim ns As New List(Of String)
            For Each n As TreeNode In node.Nodes
                Dim ns = FindReportNames(n)
                names.AddRange(ns)
            Next
        End If
        Return names
    End Function

    Public Sub ExportReportTemplate(reportNames As List(Of String), fileName As String, Optional dbFile As Boolean = True)
        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            Using conn As New OleDbConnection(CONNECTION_STRING)
                excelApp = New Excel.Application
#If DEBUG Then
                excelApp.Visible = True
#End If

                wb = excelApp.Workbooks.Open(BLANK_FILE, [ReadOnly]:=True)
                For i As Integer = 0 To reportNames.Count - 1
                    Dim sql As String = String.Format("select {0} from {1} where {2} = '{3}' order by {4}", _
                                                      ReportMasterTable.COLUMN_NAME, ReportMasterTable.TABLE_NAME, _
                                                      ReportMasterTable.REPORT_NAME, reportNames(i), ReportMasterTable.ORDER)
                    Dim cmd As New OleDbCommand(sql, conn)
                    Dim adapter As New OleDbDataAdapter(cmd)
                    Dim table As New DataTable

                    adapter.Fill(table)

                    Dim sht As Excel.Worksheet

                    If i = 0 Then
                        sht = wb.Worksheets(1)
                    Else
                        sht = wb.Worksheets.Add
                    End If

                    With sht
                        .Name = reportNames(i)

                        .Cells(FormUtils.REPORT_NAME_ROW, START_COLUMN).value = .Name

                        Dim endColumn As Integer = START_COLUMN + table.Rows.Count

                        Dim r As Excel.Range = .Range(.Cells(FormUtils.REPORT_NAME_ROW, START_COLUMN), .Cells(FormUtils.REPORT_NAME_ROW, endColumn))
                        r.Merge()
                        r.HorizontalAlignment = Excel.Constants.xlCenter
                        r.Font.Bold = True
                        r.Font.Size = REPORT_NAME_FONT_SIZE

                        Dim bidPriceColumn As Integer
                        Dim offerPriceColumn As Integer

                        For j As Integer = 0 To table.Rows.Count
                            Dim curColumn As Integer = START_COLUMN + j
                            r = .Cells(HEAD_ROW, curColumn)

                            If j = 0 Then
                                r.Value = NO_TEXT
                            Else
                                r.Value = table.Rows(j - 1)(0)
                            End If
                            r.Font.Bold = True
                            r.HorizontalAlignment = Excel.Constants.xlCenter

                            If r.Value = BID_PRICE_TEXT Then
                                bidPriceColumn = curColumn
                            ElseIf r.Value = OFFER_PRICE_TEXT Then
                                offerPriceColumn = curColumn
                            End If

                            r = .Columns(curColumn)
                            r.EntireColumn.AutoFit()
                        Next

                        r = .Cells(MARKER_ROW, START_COLUMN)
                        r.Value = MARKER_TEXT1
                        r.HorizontalAlignment = Excel.Constants.xlLeft

                        r = .Cells(MARKER_ROW, endColumn)
                        r.Value = MARKER_TEXT2
                        r.HorizontalAlignment = Excel.Constants.xlRight

                        r = .Cells(SUMMARY_ROW, START_COLUMN + 1)
                        r.Value = SUMMARY_TEXT

                        AddSumFormula(.Cells(SUMMARY_ROW, bidPriceColumn))
                        AddSumFormula(.Cells(SUMMARY_ROW, offerPriceColumn))

                        r = .Range(.Cells(HEAD_ROW, START_COLUMN), .Cells(SUMMARY_ROW, endColumn))
                        FormUtils.DrawGrid(r)
                    End With
                Next
            End Using

            wb.SaveAs(fileName)
            wb.Close()
            excelApp.Quit()

            If dbFile Then
                With My.Computer.FileSystem
                    .CopyFile(DB_FILE, .GetParentPath(fileName) + "\" + .GetName(DB_FILE), True)
                End With
            End If
        Catch ex As Exception
            Log.WriteLine(ex)

            If Not IsNothing(wb) Then
                wb.Close(False)
            End If

            If Not IsNothing(excelApp) Then
                excelApp.Quit()
            End If

            ShowErrorMessage(ex.Message)
        Finally
            wb = Nothing
            excelApp = Nothing
        End Try
    End Sub
    Public Sub DrawGrid(r As Excel.Range)
        Try
            Dim styles() As Excel.XlBordersIndex
            'Dim styles() As XlBordersIndex = {XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeRight, XlBordersIndex.xlEdgeTop, XlBordersIndex.xlEdgeBottom, XlBordersIndex.xlInsideVertical, XlBordersIndex.xlInsideHorizontal}
            If r.Rows.Count > 1 Then
                styles = {Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlInsideVertical, Excel.XlBordersIndex.xlInsideHorizontal}
            Else
                styles = {Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlInsideVertical}
            End If

            For Each s As Excel.XlBordersIndex In styles
                r.Borders(s).LineStyle = Excel.XlLineStyle.xlContinuous
                r.Borders(s).Weight = Excel.XlBorderWeight.xlThin
            Next

        Catch ex As Exception
            Log.WriteLine(ex)

        End Try
    End Sub

    Public Function ReportYear() As String
        Return ImportForm.DateTimePicker1.Value.Year
    End Function

    Public Function ProjectName() As String
        Dim pname As String = ImportForm.TextBox1.Text.Trim

        If pname.Trim = "" Then
            pname = MainForm.TreeView1.SelectedNode.FullPath.Split("\")(3)
        End If

        Return pname
    End Function

    'Public Function ComboBoxContainsText(cb As ComboBox, text As String) As Boolean
    '    Dim conains As Boolean = False

    '    For Each item As ComboBox.ObjectCollection In cb.Items
    '        If item.ToString = text Then
    '            conains = True
    '            Exit For
    '        End If
    '    Next

    '    Return conains
    'End Function
End Module
