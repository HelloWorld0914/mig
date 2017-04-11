Imports CMN
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class OperateExcel
    Implements IDisposable

#Region "メンバ変数"

    ''' <summary>
    ''' Excelアプリケーション
    ''' </summary>
    Private xlApp As Excel.Application = Nothing

    ''' <summary>
    ''' Excelブック(複)
    ''' </summary>
    Private xlBooks As Excel.Workbooks = Nothing
    ''' <summary>
    ''' Excelブック(単)
    ''' </summary>
    Private xlBook As Excel.Workbook = Nothing

    ''' <summary>
    ''' Excelシート(複)
    ''' </summary>
    Private xlSheets As Excel.Sheets = Nothing
    ''' <summary>
    ''' Excelシート(単)
    ''' </summary>
    Private xlSheet As Excel.Worksheet = Nothing

    ''' <summary>
    ''' Excelファイル名(フルパス)
    ''' </summary>
    Private xlFileName As String = Nothing

    ''' <summary>
    ''' 報告書各項目の情報を格納します。(項目名、項目の列番号(頭)、文字列数)
    ''' </summary>
    Private reportColumnData As Dictionary(Of String, List(Of Integer)) = Nothing

    ''' <summary>
    ''' エラーメッセージ作成用
    ''' </summary>
    ''' <remarks>
    ''' 日付        作成・変更者   内容
    ''' 2014.11.04  NESI           初版
    ''' </remarks>
    Private mSysErrMes As CMN.SysErrorMakeMessage = Nothing

#End Region

#Region "DB"

    Private mSqlConn As SqlConnection = Nothing
    Private mSqlAdpt As SqlDataAdapter = Nothing

    Private mOpenStateFlg As Boolean = False

    'Private Sub Open()
    '    If mOpenStateFlg = False Then
    '        mSqlConn = New SqlConnection("Data Source=192.168.10.8\sqlsvr01;Initial Catalog=移行DB;User Id=sa;Password=Nesi-2224")
    '        mSqlConn.Open()

    '        mOpenStateFlg = True
    '    End If
    'End Sub
    'Private Sub Close()
    '    mSqlConn.Close()
    '    mSqlConn.Dispose()
    '    mSqlConn = Nothing
    '    mOpenStateFlg = False
    'End Sub

    ''' <summary>
    ''' 項目情報を取得します。
    ''' </summary>
    ''' <param name="report">報告書種類。</param>
    ''' <param name="reportIn">報告書入力種類。</param>
    ''' <returns></returns>
    Private Function GetColumnData(ByVal report As ReportType, ByVal reportIn As ReportInputType) As DataTable
        Try
            Dim dTable As New DataTable
            Call Open()

            Dim cmd As New SqlCommand
            cmd.Connection = mSqlConn
            Dim sql As New StringBuilder
            With sql
                .Append("select * ")
                .Append("from ")
                .Append("[移行DB].[dbo].[MIG_" & report.ToString & "_項目情報_" & reportIn.ToString & "]")
            End With
            cmd.CommandText = sql.ToString

            Using sqlAdpt = New SqlDataAdapter
                sqlAdpt.SelectCommand = cmd
                sqlAdpt.Fill(dTable)
            End Using

            Return dTable
        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' サーバ上のテーブルをクローンします。
    ''' </summary>
    ''' <param name="report">報告書の種類。</param>
    ''' <remarks>
    ''' 日付        作成・変更者   内容
    ''' 2017.03.30  NESI           初版
    ''' </remarks>
    Private Function CloneTable(ByVal report As ReportType, ByVal reportIn As ReportInputType) As DataTable
        Try
            Dim dTable As New DataTable
            Call Open()

            Dim cmd As New SqlCommand
            cmd.Connection = mSqlConn
            cmd.CommandText = "select * from [移行DB].[dbo].[MIG_" & report.ToString & "_報告書_" & reportIn.ToString & "]"

            mSqlAdpt = New SqlDataAdapter
            mSqlAdpt.SelectCommand = cmd
            mSqlAdpt.Fill(dTable)

            Return dTable

        Catch ex As Exception
            Throw
        End Try
    End Function




    ''' <summary>
    ''' テーブルの変更を反映します。
    ''' </summary>
    ''' <param name="dTable"></param>
    Private Sub PushTable(ByVal dTable As DataTable)
        Try
            Dim cmdBuilder As New SqlCommandBuilder(mSqlAdpt)
            mSqlAdpt.Update(dTable)

            mSqlAdpt.Dispose()
            mSqlAdpt = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private dbObj As CMN.DbAccess
    Private Function Open() As Boolean
        Try
            dbObj = New CMN.DbAccess
            dbObj.SetCheckDBNames("移行DB")
            If dbObj.Open("192.168.10.8\sqlsvr01", "sa", "Techno38#") <> CMN.DbAccess.DB_RESULT.OK Then
                mSysErrMes.AddMes(Reflection.MethodBase.GetCurrentMethod.DeclaringType.ToString,
                                 Reflection.MethodBase.GetCurrentMethod.Name, dbObj.ErrMessage)
                Return False
            End If

            Return True

        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Sub Close()
        Try
            If Not dbObj Is Nothing Then
                If dbObj.IsOpen Then
                    dbObj.Close()
                End If
                dbObj.Dispose()
                dbObj = Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Function CloneTable(ByVal sql As String) As DataTable
        Try
            Dim dTable As New DataTable
            Call Open()

            Using cmd As New SqlCommand(sql, dbObj.ConnectObj)
                mSqlAdpt = New SqlDataAdapter
                mSqlAdpt.SelectCommand = cmd
                mSqlAdpt.Fill(dTable)
            End Using

            Return dTable

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' SQLを実行し、結果を返します。
    ''' </summary>
    ''' <param name="sql">SQL分。</param>
    ''' <returns></returns>
    Private Function GetDBData(ByVal sql As String) As DataTable
        Try
            Dim dTable As New DataTable
            Call Open()

            Using cmd As New SqlCommand(sql, dbObj.ConnectObj)
                Using sqlAdpt As New SqlDataAdapter
                    sqlAdpt.SelectCommand = cmd
                    sqlAdpt.Fill(dTable)
                End Using
            End Using

            Return dTable

        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

    ''' <summary>
    ''' ReleaseExcelComObjectメソッドの引数用
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum EnumReleaseType
        Sheet
        Sheets
        Book
        WorkBooks
        App
    End Enum

    Private Enum ReportType
        ICR
        PIL
    End Enum
    Private Enum ReportInputType
        ヘッダ
        エントリー
    End Enum

    ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks>
        ''' Excelアプリケーションを起動する
        ''' </remarks>
    Public Sub New()

        'Excel起動
        xlApp = New Excel.Application()

        ' アラートメッセージの表示／非表示を設定
        xlApp.DisplayAlerts = False   ' 非表示
        ' Excelの表示／非表示を設定
        xlApp.Visible = False   ' 非表示

    End Sub

    ''' <summary>
    ''' 報告書のデータを読み込みます。(ICR、PIL)
    ''' </summary>
    ''' <param name="filePath">報告書のフルパス。</param>
    Public Sub LoadReport(ByVal filePath As String)
        Try
            Dim rep As ReportType =
                If(Path.GetFileNameWithoutExtension(filePath) Like "*ICR*", ReportType.ICR, ReportType.PIL)

            Call OpenExcel(filePath)

            Call GetReportHedding(rep)
            Call GetReportEntry(rep)

            MessageBox.Show("成功")

        Catch ex As Exception
            Throw
        Finally
            Me.Dispose()
        End Try
    End Sub



    ''' <summary>
    ''' 指定したExcelファイルを開きます。
    ''' </summary>
    ''' <param name="filePath">Excelファイルのフルパス。</param>
    Private Sub OpenExcel(ByVal filePath As String)
        Try
            '作成するExcelファイル名
            xlFileName = filePath

            'Excelファイルを開く
            xlBooks = xlApp.Workbooks
            xlBook = xlBooks.Open(xlFileName)

            'Excelファイルのシートを開く
            xlSheets = xlBook.Worksheets
            xlSheet = xlSheets(2)

        Catch ex As Exception
            Throw

        End Try

    End Sub

    ''' <summary>
    ''' 報告書ヘッダ情報を取得します。
    ''' </summary>
    ''' <param name="rep"></param>
    Private Sub GetReportHedding(ByVal rep As ReportType)
        Try
            Dim dTable As New DataTable
            dTable = CloneTable(rep, ReportInputType.ヘッダ)

            '最初の行番号
            Dim currentRow As Integer = 11
            '項目の情報取得
            Dim columnData As New DataTable
            columnData = GetColumnData(rep, ReportInputType.ヘッダ)

            Dim reportData As New Dictionary(Of String, String)
            For Each row As DataRow In columnData.Rows
                Dim arrData(,) As Object = xlSheet.Range(xlSheet.Cells(currentRow, row("列番号")), xlSheet.Cells(currentRow, row("列番号") + row("文字数") - 1)).Value
                reportData.Add(row("列名"), CombineCharacter(arrData))
            Next

            dTable.Rows.Add(GetFormatReportData(dTable, reportData))

            Call PushTable(dTable)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 取得した報告書データを格納可能な形式に整形します。
    ''' </summary>
    ''' <param name="dTable">格納元DataTable。</param>
    ''' <param name="reportData">整形対象のデータ。</param>
    ''' <returns></returns>
    Private Function GetFormatReportData(ByVal dTable As DataTable, reportData As Dictionary(Of String, String)) As DataRow
        Try
            Dim dRow As DataRow = dTable.NewRow
            For Each column As String In reportData.Keys
                Select Case column
                    Case "報告期間FROM", "報告期間TO", "在庫変動年月日"
                        dRow(column) = DateTime.ParseExact("20" & reportData(column), "yyyyMMdd", Nothing)
                    Case Else
                        dRow(column) = reportData(column)
                End Select
            Next
            Return dRow
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub GetReportEntry(ByVal rep As ReportType)
        Try
            Dim dTable As New DataTable
            dTable = CloneTable(rep, ReportInputType.ヘッダ)

            '最初の行番号
            Dim currentRow As Integer = 20
            '項目の情報取得
            Dim columnData As New DataTable
            columnData = GetColumnData(rep, ReportInputType.エントリー)

            Do While (Not IsNothing(xlSheet.Range("B" & currentRow).Value))
                'ページ終端行ならスキップ
                If Not xlSheet.Range("B" & currentRow).Value.ToString = "1" Then
                    Dim reportData As New Dictionary(Of String, String)
                    For Each row As DataRow In columnData.Rows
                        Dim arrData(,) As Object = xlSheet.Range(xlSheet.Cells(currentRow, row("列番号")), xlSheet.Cells(currentRow, row("列番号") + row("文字数") - 1)).Value
                        reportData.Add(row("列名"), CombineCharacter(arrData))
                    Next
                    dTable.Rows.Add(GetFormatReportData(dTable, reportData))

                End If

                '次の行の分カウント増加
                currentRow = currentRow + 2

            Loop

            Call PushTable(dTable)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Function CombineCharacter(ByVal target As Object(,)) As String
        Dim combinedString As String = ""
        For Each v As String In target
            combinedString = combinedString + v
        Next
        Return combinedString
    End Function



#Region "dispose"

    ''' <summary>
    ''' デストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        MyBase.Finalize()

        ' Disposeが呼ばれてなかったら呼び出す
        If disposedValue = False Then
            Me.Dispose()
        End If
    End Sub

    Private disposedValue As Boolean = False        ' 重複する呼び出しを検出するには

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ReleaseXlsComObject()
            End If
        End If
        Me.disposedValue = True
    End Sub

    ' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    ''' <summary>
    ''' COMオブジェクトの解放処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReleaseXlsComObject()
        Try
            ' xlSheet解放
            If Not xlSheet Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet)
                xlSheet = Nothing
            End If

            ' xlWorkSheets解放
            If Not xlSheets Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets)
                xlSheets = Nothing
            End If

            ' xlBook解放
            If Not xlBook Is Nothing Then
                Try
                    xlBook.Close()
                Finally
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook)
                    xlBook = Nothing
                End Try
            End If

            ' xlBooks解放
            If Not xlBooks Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks)
                xlBooks = Nothing
            End If

            ' xlApp解放
            If Not xlApp Is Nothing Then
                Try
                    ' アラートを戻す
                    xlApp.DisplayAlerts = True
                    xlApp.Quit()
                Finally
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
                    xlApp = Nothing
                End Try
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub

    '''' <summary>
    '''' COMオブジェクトの解放処理
    '''' </summary>
    '''' <remarks></remarks>
    'Private Sub ReleaseXlsComObject()
    '    Try
    '        ' xlSheet解放
    '        If Not xlSheet Is Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet)
    '            xlSheet = Nothing
    '        End If

    '        'If ReleaseType = EnumReleaseType.Sheet Then
    '        '    Exit Sub
    '        'End If

    '        ' xlWorkSheets解放
    '        If Not xlSheets Is Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets)
    '            xlSheets = Nothing
    '        End If

    '        'If ReleaseType = EnumReleaseType.Sheets Then
    '        '    Exit Sub
    '        'End If

    '        ' xlBook解放
    '        If Not xlBook Is Nothing Then
    '            Try
    '                xlBook.Close()
    '            Finally
    '                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook)
    '                xlBook = Nothing
    '            End Try
    '        End If

    '        'If ReleaseType = EnumReleaseType.Book Then
    '        '    Exit Sub
    '        'End If

    '        ' xlBooks解放
    '        If Not xlBooks Is Nothing Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks)
    '            xlBooks = Nothing
    '        End If

    '        'If ReleaseType = EnumReleaseType.WorkBooks Then
    '        '    Exit Sub
    '        'End If

    '        ' xlApp解放
    '        If Not xlApp Is Nothing Then
    '            Try
    '                ' アラートを戻す
    '                xlApp.DisplayAlerts = True
    '                xlApp.Quit()
    '            Finally
    '                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
    '                xlApp = Nothing
    '            End Try
    '        End If

    '    Catch ex As Exception
    '        Throw
    '    End Try

    'End Sub
#End Region


End Class

