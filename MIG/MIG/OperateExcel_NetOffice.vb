Imports System.Data
Imports System.Text

Public Class OperateExcel_NetOffice
    Implements IDisposable

#Region "メンバ変数"
    ''' <summary>
    ''' Excelアプリケーション。(アンマネージリソースのため取扱い注意)
    ''' </summary>
    Private mExcelApp As NetOffice.ExcelApi.Application
    ''' <summary>
    ''' フォーマット用Excelブック。(アンマネージリソースのため取扱い注意)
    ''' </summary>
    Private mExcelBookForFormat As NetOffice.ExcelApi.Workbook

    ''' <summary>
    ''' 出力用フォーマットのフルパス。
    ''' </summary>
    Private mFormatPath As String = "D:\Git\MIG\MIG\MIG\報告データ一覧.xlsx"

    Private mErrorFilePath As String = ""

    Private mHeaderDataList As List(Of List(Of String)) = Nothing
    Private mICRDataList As List(Of List(Of String)) = Nothing
    Private mPILDataList As List(Of List(Of String)) = Nothing
    Private mMBRDataList As List(Of List(Of String)) = Nothing
    Private mOCR1DataList As List(Of List(Of String)) = Nothing
    Private mOCR3DataList As List(Of List(Of String)) = Nothing
    Private mConciseDataList As List(Of List(Of String)) = Nothing

    Private mHeaderFieldList As List(Of String) = Nothing
    Private mICRFieldList As List(Of String) = Nothing
    Private mPILFieldList As List(Of String) = Nothing
    Private mMBRFieldList As List(Of String) = Nothing
    Private mOCR1FieldList As List(Of String) = Nothing
    Private mOCR3FieldList As List(Of String) = Nothing
    Private mConciseFieldList As List(Of String) = Nothing

#End Region

#Region "列挙型"
    ''' <summary>
    ''' 報告書の種類を表します。
    ''' </summary>
    Private Enum ReportType
        ICR = 3
        PIL = 4
        MBR = 5
        OCR1 = 6
        OCR3 = 7
        注釈 = 8
    End Enum
    ''' <summary>
    ''' 報告書の入力種類を表します。
    ''' </summary>
    Private Enum ReportInput
        Header = 0
        Entry = 1
    End Enum
    ''' <summary>
    ''' 報告書が作成された年代を表します。
    ''' </summary>
    Private Enum ReportYear
        Before2012 = 2
        After2013 = 4
    End Enum

#End Region

#Region "メイン"
    ''' <summary>
    ''' 指定したフォルダ内の全ファイルを、報告書形式からリスト形式に移行します。
    ''' </summary>
    ''' <param name="folderPath">対象フォルダのフルパス。</param>
    Public Sub MigrateReportData(ByVal folderPath As String)
        Try
            Dim sw As New System.Diagnostics.Stopwatch()
            sw.Start()

            Dim dTableForHeader As New DataTable
            Dim dTableDictionalyForEntry As New Dictionary(Of String, DataTable)

            Call openExcelBook()
            Call InitDataTable(dTableForHeader, dTableDictionalyForEntry)

            Call loadReport(dTableForHeader, dTableDictionalyForEntry, getFilePath(folderPath))

            Call OutputReportData(dTableForHeader, dTableDictionalyForEntry)

            sw.Stop()

            MessageBox.Show("Success. " & sw.Elapsed.ToString)

        Catch ex As Exception
            MessageBox.Show(ex.Message & vbCrLf & "Path : " & mErrorFilePath)
        Finally
            Dispose()
            MessageBox.Show("Dispose Completed.")
        End Try
    End Sub
#End Region

#Region "報告書取込"
    ''' <summary>
    ''' 各種DataTableの初期設定を行います。
    ''' </summary>
    ''' <param name="dTableForHeader">ヘッダ用DataTable。</param>
    ''' <param name="dTableDictionaryForEntry">エントリー用DataTable。</param>
    Private Sub InitDataTable(ByRef dTableForHeader As DataTable, ByRef dTableDictionaryForEntry As Dictionary(Of String, DataTable))
        Try
            dTableForHeader = formatDataTable()

            Dim reportTypeList As New List(Of ReportType)(New ReportType() {ReportType.ICR, ReportType.PIL, ReportType.MBR, ReportType.OCR1, ReportType.OCR3, ReportType.注釈})
            For Each rep As ReportType In reportTypeList
                dTableDictionaryForEntry.Add(rep.ToString, formatDataTable(rep))
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub setFieldList(Optional ByVal repType As ReportType = 0)
        Try
            If repType = ReportType.ICR Then
                If IsNothing(mHeaderFieldList) Then
                    mHeaderFieldList = New List(Of String)
                End If
            ElseIf repType = ReportType.PIL Then

            ElseIf repType = ReportType.MBR Then

            ElseIf repType = ReportType.OCR1 Then

            ElseIf repType = ReportType.OCR3 Then

            ElseIf repType = ReportType.注釈 Then

            Else

            End If


        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 指定した報告書を読み込み、データを取得します。
    ''' </summary>
    ''' <param name="dTableForHeader">データ格納用DataTable(ヘッダ)。</param>
    ''' <param name="dTableDictionalyForEntry">データ格納用Dictionary(エントリー)。</param>
    ''' <param name="filePathList">対象ファイルのフルパスを格納したリスト。</param>
    Private Sub loadReport(ByRef dTableForHeader As DataTable, ByRef dTableDictionalyForEntry As Dictionary(Of String, DataTable), ByVal filePathList As List(Of String))
        Try
            '全ファイルやるやつ
            For Each filePath As String In filePathList
                mErrorFilePath = filePath
                'Excelブックオープン
                Dim xlBook As NetOffice.ExcelApi.Workbook = openExcelBook(filePath)
                '2012年以前か2012年以降かを判断する
                If IO.Path.GetFileNameWithoutExtension(filePath) Like "*計量管理報告書*" OrElse
                    IO.Path.GetFileNameWithoutExtension(filePath) Like "*ICR*" AndAlso IO.Path.GetFileNameWithoutExtension(filePath) Like "*OCR1*" Then
                    '2012年以前
                    Call loadReportDataForBefore2012(xlBook, dTableForHeader, dTableDictionalyForEntry)
                Else
                    '2013年以降
                    Call loadReportDataForAfter2013(xlBook, dTableForHeader, dTableDictionalyForEntry)
                End If

            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 2012年以前の報告書を読み込みます。
    ''' </summary>
    ''' <param name="xlBook">Excelブック。</param>
    ''' <param name="dTableHeader">ヘッダ用DataTable。</param>
    ''' <param name="dTableDictionaryForEntry">エントリー用Dictionary。</param>
    Private Sub loadReportDataForBefore2012(ByVal xlBook As NetOffice.ExcelApi.Workbook, ByRef dTableHeader As DataTable, ByRef dTableDictionaryForEntry As Dictionary(Of String, DataTable))
        Try
            '全シート名取得
            Dim sheetNameList As List(Of String) = getSheetName(xlBook)
            For Each sheetName As String In sheetNameList
                'Excelシートオープン
                Dim xlSheet As NetOffice.ExcelApi.Worksheet = getExcelSheet(sheetName, xlBook)
                '報告書種類取得
                Dim rep As ReportType = getReportType(sheetName)
                'ヘッダ情報取得
                Call setReportHeader(dTableHeader, xlSheet, rep, ReportYear.Before2012)
                'エントリー情報取得
                Call setReportEntry(dTableDictionaryForEntry(rep.ToString), xlSheet, rep, ReportYear.Before2012)
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 2013年以降の報告書を読み込みます。
    ''' </summary>
    ''' <param name="xlBook">Excelブック。</param>
    ''' <param name="dTableHeader">ヘッダ用DataTable。</param>
    ''' <param name="dTableDictionaryForEntry">エントリー用Dictionary。</param>
    Private Sub loadReportDataForAfter2013(ByVal xlBook As NetOffice.ExcelApi.Workbook, ByRef dTableHeader As DataTable, ByRef dTableDictionaryForEntry As Dictionary(Of String, DataTable))
        Try
            '全シート名取得
            Dim sheetNameList As List(Of String) = getSheetName(xlBook)
            '対象シート名取得
            Dim targetSheetName As String = ""
            Dim conciseNoteName As String = ""
            For Each sheetName As String In sheetNameList
                '[〇〇入力]というシート名が対象
                If sheetName Like "*入力*" Then targetSheetName = sheetName
                '注釈があればTrue
                If sheetName Like "*注釈*" Then conciseNoteName = sheetName
            Next

            'Excelシートオープン
            Dim xlSheet As NetOffice.ExcelApi.Worksheet = getExcelSheet(targetSheetName, xlBook)
            '報告書種類取得
            Dim rep As ReportType = getReportType(xlBook.FullName)
            'ヘッダ情報取得
            Call setReportHeader(dTableHeader, xlSheet, rep, ReportYear.After2013)
            'エントリー情報取得
            Call setReportEntry(dTableDictionaryForEntry(rep.ToString), xlSheet, rep, ReportYear.After2013)
            '注釈があれば追加する
            If Not conciseNoteName = "" Then
                Dim xlSheetForConciseNote As NetOffice.ExcelApi.Worksheet = getExcelSheet(conciseNoteName, xlBook)
                Call setReportEntry(dTableDictionaryForEntry(ReportType.注釈.ToString), xlSheetForConciseNote, ReportType.注釈, ReportYear.After2013)
            End If


        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' ヘッダ情報を取得してやろう！
    ''' </summary>
    ''' <param name="dTableHeader"></param>
    ''' <param name="xlSheet"></param>
    ''' <param name="repType"></param>
    ''' <param name="repYear"></param>
    Private Sub setReportHeader(ByVal dTableHeader As DataTable, ByVal xlSheet As NetOffice.ExcelApi.Worksheet, ByVal repType As ReportType, ByVal repYear As ReportYear)
        Try
            Dim targetRowIdx As Integer = 0 '対象の行インデックス
            Dim startColumnIdx As Integer = 0 '対象の最初の列インデックス
            Dim endColumnIdx As Integer = 0　'対象の最後の列インデックス

            '各種位置を年代別の報告書に合わせる
            If repYear = ReportYear.Before2012 Then
                '2012年以前
                targetRowIdx = 4
                startColumnIdx = 1
                endColumnIdx = 1
            Else
                '2013年以降
                If repType = ReportType.ICR OrElse repType = ReportType.OCR1 OrElse repType = ReportType.OCR3 OrElse repType = ReportType.注釈 Then
                    targetRowIdx = 11
                ElseIf repType = ReportType.PIL Then
                    targetRowIdx = 12
                ElseIf repType = ReportType.MBR Then
                    targetRowIdx = 13
                End If
                startColumnIdx = 2
                endColumnIdx = 2
            End If

            '項目の文字数を取得
            Dim fieldLengthList As List(Of String) = getFieldLengthList(repType, ReportInput.Header, repYear)
            '最後の列インデックスを割り出す
            For Each length As String In fieldLengthList
                endColumnIdx = endColumnIdx + CInt(length)
            Next

            'ヘッダ情報を取得
            Dim excelDataList As List(Of String) = getRangeExcelData(xlSheet, targetRowIdx, startColumnIdx, targetRowIdx, endColumnIdx)

            'ヘッダ用DataTableに格納
            Call setLoadedReportData(dTableHeader, excelDataList, fieldLengthList, repType, ReportInput.Header)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' エントリー情報を取得してやろう！
    ''' </summary>
    ''' <param name="dTableEntry"></param>
    ''' <param name="xlSheet"></param>
    ''' <param name="repType"></param>
    ''' <param name="repYear"></param>
    Private Sub setReportEntry(ByVal xlSheet As NetOffice.ExcelApi.Worksheet, ByVal repType As ReportType, ByVal repYear As ReportYear)
        Try
            Dim targetRowIdx As Integer = 0 '対象の行インデックス
            Dim startColumnIdx As Integer = 0 '対象の最初の列インデックス
            Dim endColumnIdx As Integer = 0　'対象の最後の列インデックス
            Dim incremental As Integer = 0 '次の行に行くための増分

            '各種位置を年代別の報告書に合わせる
            If repYear = ReportYear.Before2012 Then
                '2012年以前
                targetRowIdx = 10
                startColumnIdx = 1
                endColumnIdx = 1
                incremental = 1
            Else
                '2013年以降
                If repType = ReportType.ICR OrElse repType = ReportType.PIL OrElse repType = ReportType.注釈 Then
                    targetRowIdx = 20
                ElseIf repType = ReportType.MBR Then
                    targetRowIdx = 21
                ElseIf repType = ReportType.OCR1 OrElse repType = ReportType.OCR3 Then
                    targetRowIdx = 19
                End If
                startColumnIdx = 2
                endColumnIdx = 2
                incremental = 2
            End If

            '項目の文字数を取得
            Dim fieldLengthList As List(Of String) = getFieldLengthList(repType, ReportInput.Entry, repYear)
            '最後の列インデックスを割り出す
            For Each length As String In fieldLengthList
                endColumnIdx = endColumnIdx + CInt(length)
            Next

            Do While (Not IsNothing(xlSheet.Cells(targetRowIdx, startColumnIdx).Value))
                'ページ終端行ならスキップ
                If Not xlSheet.Cells(targetRowIdx, startColumnIdx).Value.ToString = "1" Then
                    'ヘッダ情報を取得
                    Dim excelDataList As List(Of String) = getRangeExcelData(xlSheet, targetRowIdx, startColumnIdx, targetRowIdx, endColumnIdx)
                    'ヘッダ用DataTableに格納
                    Call setLoadedReportData(dTableEntry, excelDataList, fieldLengthList, repType, ReportInput.Entry)
                End If
                '次の行の分カウント増加
                targetRowIdx = targetRowIdx + incremental
            Loop

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 取得した報告書データを指定したDataTableに追加します。
    ''' </summary>
    ''' <param name="dTable">データを追加するDataTable。</param>
    ''' <param name="loadedDataList">取得したデータのリスト。</param>
    ''' <param name="fieldLengthList">項目別文字数のリスト。</param>
    ''' <param name="repType">報告書の種類。</param>
    ''' <param name="repInput">報告書の入力種類。</param>
    Private Sub setLoadedReportData(ByRef dTable As DataTable, ByVal loadedDataList As List(Of String), ByVal fieldLengthList As List(Of String), ByVal repType As ReportType, ByVal repInput As ReportInput)
        Try
            '報告書データリスト
            Dim reportDataList As New List(Of String)
            '取得したデータと項目別文字数リストを基に、報告書データを作成する
            For Each fieldLength As Integer In fieldLengthList
                Dim combined As New StringBuilder
                For i As Integer = 0 To fieldLength - 1
                    combined.Append(loadedDataList(0))
                    loadedDataList.RemoveAt(0)　'使用後は頭から削除
                Next
                reportDataList.Add(combined.ToString)
            Next
            If repInput = ReportInput.Header Then
                'ヘッダ情報の場合、[事業所コード]と[扱者氏名]の情報を削除する
                reportDataList.RemoveAt(0)
                reportDataList.RemoveAt(reportDataList.Count - 1)
            ElseIf repInput = ReportInput.Entry Then
                If repType = ReportType.OCR3 Then
                    reportDataList.RemoveAt(6)
                ElseIf repType = ReportType.注釈 Then
                    reportDataList.Insert(3, "")
                End If
            End If

            '追加用DataRow
            Dim dRow As DataRow = dTable.NewRow

            '報告書データ用インデックス
            Dim reportDataIdx As Integer = 0
            '列数分回す
            For Each column As DataColumn In dRow.Table.Columns
                'ヘッダ情報の場合
                If repInput = ReportInput.Header Then
                    '項目名が[F_報告期間]かつ報告書種類が在庫系の場合、スキップする
                    If column.ColumnName = "報告期間FROM" Then
                        If repType = ReportType.PIL OrElse repType = ReportType.OCR3 Then Continue For
                    End If
                    '項目名が[報告書タイプ]の場合、報告書種類を追加して終了する
                    If column.ColumnName = "報告書区分" Then
                        Dim repCode As String = ""
                        Select Case repType
                            Case ReportType.ICR
                                repCode = "1"
                            Case ReportType.PIL
                                repCode = "4"
                            Case ReportType.MBR
                                repCode = "6"
                            Case ReportType.OCR1
                                repCode = "A"
                            Case ReportType.OCR3
                                repCode = "E"
                            Case ReportType.注釈
                                repCode = "2"
                        End Select
                        dRow(column.ColumnName) = repCode
                        Exit For
                    End If
                ElseIf repInput = ReportInput.Entry Then
                    If column.ColumnName = "棚卸日" OrElse column.ColumnName = "棚卸実施日" Then
                        Continue For
                    End If
                End If
                dRow(column.ColumnName) = reportDataList(reportDataIdx)
                reportDataIdx = reportDataIdx + 1
            Next

            'データ追加
            dTable.Rows.Add(dRow)

        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "報告書出力"
    Private Sub outputReportData()
        Try
            '新しいExcelブックをデスクトップに作成する
            Dim newFilePath As String =
                System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\報告データ_" & DateTime.Now.ToString("yyyyMMdd") & ".xlsx"
            Call copyExcelBook(mFormatPath, newFilePath, True)
            '[_項目文字数情報]シートを削除する
            Call deleteExcelSheet(newFilePath, "_項目文字数情報")

            Dim xlBook As NetOffice.ExcelApi.Workbook = openExcelBook(newFilePath)
            Dim sheetNameList As List(Of String) = getSheetName(xlBook)
            For Each sheetName As String In sheetNameList
                Dim xlSheet As NetOffice.ExcelApi.Worksheet = getExcelSheet(sheetName, xlBook)
                Dim outputList As List(Of List(Of String)) = Nothing

                Select Case getReportType(sheetName)
                    Case ReportType.ICR
                        outputList = mICRList
                    Case ReportType.PIL
                        outputList = mPILList
                    Case ReportType.MBR
                        outputList = mMBRList
                    Case ReportType.OCR1
                        outputList = mOCR1List
                    Case ReportType.OCR3
                        outputList = mOCR3List
                    Case ReportType.注釈
                        outputList = mConciseList
                    Case Else
                        outputList = mHeaderList
                End Select

                If outputList.Count = 0 Then Continue For

                Dim maxRow As Integer = outputList.Count
                Dim maxColumn As Integer = outputList(0).Count

                Dim outputArray(maxRow - 1, maxColumn - 1) As String
                For i As Integer = 0 To maxRow - 1
                    For j As Integer = 0 To maxColumn - 1
                        outputArray(i, j) = outputList(i)(j)
                    Next
                Next

                Dim startRange As NetOffice.ExcelApi.Range = xlSheet.Cells(2, 1)
                Dim endRange As NetOffice.ExcelApi.Range = xlSheet.Cells(maxRow, maxColumn)

                Dim columnList As List(Of String) = getRangeExcelData(xlSheet, 1, 1, 1, 26, True)
                For i As Integer = 0 To columnList.Count - 1
                    xlSheet.Range(xlSheet.Cells(2, i + 1), xlSheet.Cells(maxRow, i + 1)).Value = If(columnList(i) Like "*日*", "yyyy/m/d", "@")
                Next

                xlSheet.Range(startRange, endRange).Value = outputArray

            Next
            xlBook.Save()


        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Excel系ユーティリティ"
    ''' <summary>
    ''' Excelアプリケーション、フォーマット用Excelブックを定義します。
    ''' </summary>
    Private Sub openExcelBook()
        Try
            If IsNothing(mExcelApp) Then
                mExcelApp = New NetOffice.ExcelApi.Application
                mExcelApp.DisplayAlerts = False
                mExcelApp.Visible = False
            End If
            If IsNothing(mExcelBookForFormat) Then
                mExcelBookForFormat = mExcelApp.Workbooks.Open(mFormatPath)
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub
    ''' <summary>
    ''' 指定したExcelファイルを開き、ファイル情報を返します。
    ''' </summary>
    ''' <param name="filePath">対象ファイルのフルパス。</param>
    ''' <returns></returns>
    Private Function openExcelBook(ByVal filePath As String) As NetOffice.ExcelApi.Workbook
        Try
            If IsNothing(mExcelApp) Then
                mExcelApp = New NetOffice.ExcelApi.Application
                mExcelApp.DisplayAlerts = False
                mExcelApp.Visible = False
            End If
            Return mExcelApp.Workbooks.Open(filePath)

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定したExcelシートを開き、シート情報を返します。
    ''' </summary>
    ''' <param name="sheetName"></param>
    ''' <param name="xlBook"></param>
    ''' <returns></returns>
    Private Function getExcelSheet(ByVal sheetName As String, ByVal xlBook As NetOffice.ExcelApi.Workbook) As NetOffice.ExcelApi.Worksheet
        Try
            Return CType(xlBook.Worksheets(sheetName), NetOffice.ExcelApi.Worksheet)

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定したファイルの報告書種類を取得します。
    ''' </summary>
    ''' <param name="file">対象名。</param>
    ''' <returns>報告書種類。対象外のファイルの場合Nothingを返す。</returns>
    Private Function getReportType(ByVal file As String) As ReportType
        Try
            Dim rep As ReportType = 0
            If IO.Path.GetFileNameWithoutExtension(file) Like "*ICR*" Then
                rep = ReportType.ICR
            ElseIf IO.Path.GetFileNameWithoutExtension(file) Like "*PIL*" Then
                rep = ReportType.PIL
            ElseIf IO.Path.GetFileNameWithoutExtension(file) Like "*MBR*" Then
                rep = ReportType.MBR
            ElseIf IO.Path.GetFileNameWithoutExtension(file) Like "*OCR1*" Then
                rep = ReportType.OCR1
            ElseIf IO.Path.GetFileNameWithoutExtension(file) Like "*OCR3*" Then
                rep = ReportType.OCR3
            ElseIf IO.Path.GetFileNameWithoutExtension(file) Like "*注釈*" Then
                rep = ReportType.注釈
            End If

            If rep = 0 AndAlso IO.Path.GetFileNameWithoutExtension(file) Like "*OCR*" Then
                rep = ReportType.OCR1
            End If

            Return rep

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定したExcelブックに属するシート名を取得します。
    ''' </summary>
    ''' <param name="xlBook">対象のExcelブック。</param>
    ''' <returns></returns>
    Private Function getSheetName(ByVal xlBook As NetOffice.ExcelApi.Workbook) As List(Of String)
        Try
            Dim sheetNameList As New List(Of String)
            For Each xlSheet As NetOffice.ExcelApi.Worksheet In xlBook.Sheets
                If xlSheet.Name Like "*Sheet*" Then Continue For
                sheetNameList.Add(xlSheet.Name)
            Next
            Return sheetNameList

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定した範囲のデータを取得します。
    ''' </summary>
    ''' <param name="xlSheet">対象シート。</param>
    ''' <param name="startRowIdx">最初のセルの行番号。</param>
    ''' <param name="startColumnIdx">最初のセルの列番号。</param>
    ''' <param name="endRowIdx">最後のセルの行番号。</param>
    ''' <param name="endColumnIdx">最後のセルの列番号。</param>
    ''' <param name="nothingMode">Trueに設定した場合、値が入っていないセルを発見したタイミングで、取得を終了する。</param>
    ''' <returns></returns>
    Private Function getRangeExcelData(ByVal xlSheet As NetOffice.ExcelApi.Worksheet,
                                       ByVal startRowIdx As Integer, ByVal startColumnIdx As Integer,
                                       ByVal endRowIdx As Integer, ByVal endColumnIdx As Integer,
                                        Optional ByVal nothingMode As Boolean = False) As List(Of String)
        Try
            Dim excelDataList As New List(Of String)
            Dim strArray(,) As Object =
                CType(xlSheet.Range(xlSheet.Cells(startRowIdx, startColumnIdx), xlSheet.Cells(endRowIdx, endColumnIdx)).Value, Object(,))
            For Each str As String In strArray
                If nothingMode AndAlso IsNothing(str) Then Exit For
                excelDataList.Add(str)
            Next
            Return excelDataList

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' DataTableをヘッダ情報形式で初期化します。
    ''' </summary>
    ''' <returns></returns>
    Private Function formatDataTable() As DataTable
        Try
            Dim fieldNameList As List(Of String) = Nothing
            Using xlSheet As NetOffice.ExcelApi.Worksheet = getExcelSheet("ヘッダー情報", mExcelBookForFormat)
                fieldNameList = getRangeExcelData(xlSheet, 1, 1, 1, 26, True)
            End Using
            Dim dTable As New DataTable
            For Each fieldName As String In fieldNameList
                dTable.Columns.Add(fieldName)
            Next
            Return dTable

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定した報告書形式でDataTableを初期化します。
    ''' </summary>
    ''' <returns></returns>
    Private Function formatDataTable(ByVal rep As ReportType) As DataTable
        Try
            Dim sheetNameList As List(Of String) = getSheetName(mExcelBookForFormat)

            Dim fieldNameList As List(Of String) = Nothing
            For Each sheetName As String In sheetNameList
                If sheetName = rep.ToString Then
                    Using xlSheet As NetOffice.ExcelApi.Worksheet = getExcelSheet(sheetName, mExcelBookForFormat)
                        fieldNameList = getRangeExcelData(xlSheet, 1, 1, 1, 26, True)
                    End Using
                    Exit For
                End If
            Next

            Dim dTable As New DataTable
            For Each fieldName As String In fieldNameList
                dTable.Columns.Add(fieldName)
            Next

            Return dTable

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定した種類のフィールド情報を取得します。
    ''' </summary>
    ''' <param name="repType"></param>
    ''' <param name="repYear"></param>
    ''' <returns></returns>
    Private Function getFieldLengthList(ByVal repType As ReportType, ByVal repInput As ReportInput, ByVal repYear As ReportYear) As List(Of String)
        Try
            Dim xlSheet As NetOffice.ExcelApi.Worksheet = getExcelSheet("_項目文字数情報", mExcelBookForFormat)

            Dim fieldLenfthList As New List(Of String)
            Dim strField As String = CType(xlSheet.Cells(CType(repType, Short), CType(repYear, Short) + CType(repInput, Short)).Value, String)
            fieldLenfthList.AddRange(strField.ToString.Split(CType(",", Char)))

            Return fieldLenfthList

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定したシートを削除します。
    ''' </summary>
    ''' <param name="filePath">対象Excelファイルのフルパス。</param>
    ''' <param name="sheetName">対象シート名。</param>
    Private Sub deleteExcelSheet(ByVal filePath As String, ByVal sheetName As String)
        Try
            Using xlBook As NetOffice.ExcelApi.Workbook = openExcelBook(filePath)
                Using xlSheet As NetOffice.ExcelApi.Worksheet = getExcelSheet(sheetName, xlBook)
                    xlSheet.Delete()
                End Using
                xlBook.Save()
            End Using

        Catch ex As Exception
            Throw
        End Try
    End Sub
    ''' <summary>
    ''' パスを指定し、Excelファイルをコピーします。
    ''' </summary>
    ''' <param name="sourceFilePath">コピー元パス。</param>
    ''' <param name="copiedFilePath">コピー先パス。</param>
    ''' <param name="overwrite">Trueに設定した場合、同名のファイルが存在した場合に上書きします。</param>
    Private Sub copyExcelBook(ByVal sourceFilePath As String, ByVal copiedFilePath As String, Optional ByVal overwrite As Boolean = False)
        Try
            System.IO.File.Copy(sourceFilePath, copiedFilePath, overwrite)
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "ユーティリティ"
    ''' <summary>
    ''' 指定した配列内の要素を結合し、文字列を作成します。
    ''' </summary>
    ''' <param name="strArray">結合対象の配列。</param>
    ''' <returns>結合後の文字列。</returns>
    Private Function combineCharacter(ByVal strArray As Object(,)) As String
        Try
            Dim strBuilder As New StringBuilder
            For Each str As String In strArray
                strBuilder.Append(str)
            Next
            Return strBuilder.ToString

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定したリスト内の要素を結合し、文字列を作成します。
    ''' </summary>
    ''' <param name="strList">結合対象の配列。</param>
    ''' <returns>結合後の文字列。</returns>
    Private Function combineCharacter(ByVal strList As List(Of String)) As String
        Try
            Dim strBuilder As New StringBuilder
            For Each str As String In strList
                strBuilder.Append(str)
            Next
            Return strBuilder.ToString

        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 指定したフォルダ内に存在する全ファイルのフルパスを取得します。
    ''' </summary>
    ''' <param name="folderPath">対象フォルダのフルパス。</param>
    ''' <returns></returns>
    Public Function getFilePath(ByVal folderPath As String) As List(Of String)
        Try
            Dim fileNameList As New List(Of String)
            fileNameList.AddRange(System.IO.Directory.GetFiles(folderPath, "*", System.IO.SearchOption.AllDirectories))
            Return fileNameList

        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean = False ' 重複する呼び出しを検出するには

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: マネージ状態を破棄します (マネージ オブジェクト)。
            End If

            ' TODO: アンマネージ リソース (アンマネージ オブジェクト) を解放し、下の Finalize() をオーバーライドします。
            ' TODO: 大きなフィールドを null に設定します。
            If Not IsNothing(mExcelApp) Then
                mExcelApp.Quit()
                mExcelApp.Dispose()
                mExcelApp = Nothing
            End If

        End If
        disposedValue = True
    End Sub

    ' TODO: 上の Dispose(disposing As Boolean) にアンマネージ リソースを解放するコードが含まれる場合にのみ Finalize() をオーバーライドします。
    Protected Overrides Sub Finalize()
        ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(disposing As Boolean) に記述します。
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(disposing As Boolean) に記述します。
        Dispose(True)
        ' TODO: 上の Finalize() がオーバーライドされている場合は、次の行のコメントを解除してください。
        GC.SuppressFinalize(Me)
    End Sub
#End Region



End Class
