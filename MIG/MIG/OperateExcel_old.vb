Option Strict Off

Imports CMN
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Configuration

Public Class OperateExcel_old
    Implements IDisposable

#Region "メンバ変数"

    ''' <summary>
    ''' Excelアプリケーション
    ''' </summary>
    Private mExcelApp As Excel.Application = Nothing

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

    Private mSqlConn As SqlConnection = Nothing
    Private mSqlAdpt As SqlDataAdapter = Nothing

#End Region

#Region "DB関連"

    ''' <summary>
    ''' 接続を開きます。
    ''' </summary>
    ''' <returns></returns>
    Private Function OpenConnection() As Boolean
        Try
            If Not CheckConnectionOpend() Then
                'mSqlConn = New SqlConnection(ConfigurationManager.ConnectionStrings("MIG.MySettings.ConnectionString").ConnectionString)
                mSqlConn.Open()
            End If
            Return True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 接続を閉じます。
    ''' </summary>
    ''' <returns></returns>
    Private Function CloseConnection() As Boolean
        Try
            If CheckConnectionOpend() Then
                mSqlConn.Close()
                mSqlConn.Dispose()
                mSqlConn = Nothing
            End If

            Return True

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 接続が開いているか確認します。
    ''' </summary>
    ''' <returns></returns>
    Private Function CheckConnectionOpend() As Boolean
        Try
            If IsNothing(mSqlConn) Then Return False
            If mSqlConn.State <> ConnectionState.Open Then
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 指定したSQL文を実行し、結果を返します。
    ''' </summary>
    ''' <param name="sql">実行対象のSQL文。</param>
    ''' <param name="clone">Trueに設定した場合、DataTableの変更をDBへ反映させることができます。</param>
    ''' <returns></returns>
    Private Function GetDBData(ByVal sql As String, Optional ByVal clone As Boolean = False) As DataTable
        Try
            Dim dTable As New DataTable
            Call OpenConnection()

            Using cmd As New SqlCommand(sql, mSqlConn)
                If clone Then
                    'メンバ変数使用
                    mSqlAdpt = New SqlDataAdapter
                    mSqlAdpt.SelectCommand = cmd
                    mSqlAdpt.Fill(dTable)
                Else
                    'ローカル変数使用
                    Using sqlAdpt As New SqlDataAdapter
                        sqlAdpt.SelectCommand = cmd
                        sqlAdpt.Fill(dTable)
                    End Using
                End If
            End Using

            Return dTable

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' テーブルの変更をDBに反映します。
    ''' </summary>
    ''' <param name="dTable"></param>
    Private Sub Update(ByVal dTable As DataTable)
        Try
            Dim cmdBuilder As New SqlCommandBuilder(mSqlAdpt)
            mSqlAdpt.Update(dTable)

            mSqlAdpt.Dispose()
            mSqlAdpt = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub

    'Private dbObj As CMN.DbAccess
    '''' <summary>
    '''' サーバへの接続をオープンします。
    '''' </summary>
    '''' <returns></returns>
    'Private Function Open() As Boolean
    '    Try
    '        dbObj = New CMN.DbAccess
    '        dbObj.SetCheckDBNames("移行DB")
    '        If dbObj.Open("192.168.10.8\sqlsvr01", "sa", "Nesi-2224") <> CMN.DbAccess.DB_RESULT.OK Then
    '            mSysErrMes.AddMes(Reflection.MethodBase.GetCurrentMethod.DeclaringType.ToString,
    '                             Reflection.MethodBase.GetCurrentMethod.Name, dbObj.ErrMessage)
    '            Return False
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function
    '''' <summary>
    '''' サーバとの接続をクローズします。
    '''' </summary>
    'Private Sub Close()
    '    Try
    '        If Not dbObj Is Nothing Then
    '            If dbObj.IsOpen Then
    '                dbObj.Close()
    '            End If
    '            dbObj.Dispose()
    '            dbObj = Nothing
    '        End If
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Sub

#End Region

#Region "Enum"
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
#End Region

#Region "コンストラクタ"
    ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        ''' <remarks>
        ''' Excelアプリケーションを起動する
        ''' </remarks>
    Public Sub New()

    End Sub
#End Region

    Public Sub MigrateReportData(ByVal folderName As String)
        Try
            Dim fileNameList As New List(Of String)
            fileNameList = GetFileName(folderName)

            Call OpenExcel()

            For Each fileName As String In fileNameList
                xlBook = xlBooks.Open(xlFileName)
                xlSheets = xlBook.Worksheets
                xlSheet = xlSheets(2)




            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub


    Private Sub SetReportData()
        Try


        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Function GetFileName(ByVal folderName) As List(Of String)
        Try
            Dim fileNameList As New List(Of String)
            fileNameList.AddRange(System.IO.Directory.GetFiles(folderName, "*", System.IO.SearchOption.AllDirectories))

            Return fileNameList

        Catch ex As Exception
            Throw
        End Try
    End Function

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
    ''' Excelファイルを開きます。
    ''' </summary>
    Private Sub OpenExcel()
        Try
            ' Excel起動
            mExcelApp = New Excel.Application()
            ' アラートメッセージの表示／非表示を設定
            mExcelApp.DisplayAlerts = False   ' 非表示
            ' Excelの表示／非表示を設定
            mExcelApp.Visible = False   ' 非表示
            ' Excelファイルを開く
            xlBooks = mExcelApp.Workbooks

        Catch ex As Exception
            Throw
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
            Dim sqlBuilder As New StringBuilder
            With sqlBuilder
                .Append("select * ")
                .Append("from ")
                .Append("[移行DB].[dbo].[MIG_" & rep.ToString & "_報告書_" & ReportInputType.ヘッダ.ToString & "]")
            End With
            '更新用DataTable
            dTable = GetDBData(sqlBuilder.ToString, True)

            '最初の行番号
            Dim currentRow As Integer = 11
            '項目の情報取得
            Dim columnData As New DataTable
            sqlBuilder.Clear()
            With sqlBuilder
                .Append("select * ")
                .Append("from ")
                .Append("[移行DB].[dbo].[MIG_" & rep.ToString & "_項目情報_" & ReportInputType.ヘッダ.ToString & "]")
            End With
            columnData = GetDBData(sqlBuilder.ToString)

            Dim reportData As New Dictionary(Of String, String)
            For Each row As DataRow In columnData.Rows
                Dim arrData(,) As Object = xlSheet.Range(xlSheet.Cells(currentRow, row("列番号")), xlSheet.Cells(currentRow, row("列番号") + row("文字数") - 1)).Value
                reportData.Add(row("列名"), CombineCharacter(arrData))
            Next

            dTable.Rows.Add(GetFormatReportData(dTable, reportData))

            Call Update(dTable)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub GetReportEntry(ByVal rep As ReportType)
        Try
            Dim dTable As New DataTable
            Dim sqlBuilder As New StringBuilder
            With sqlBuilder
                .Append("select * ")
                .Append("from ")
                .Append("[移行DB].[dbo].[MIG_" & rep.ToString & "_報告書_" & ReportInputType.エントリー.ToString & "]")
            End With
            '更新用DataTable
            dTable = GetDBData(sqlBuilder.ToString, True)

            '最初の行番号
            Dim currentRow As Integer = 20
            '項目の情報取得
            Dim columnData As New DataTable
            sqlBuilder.Clear()
            With sqlBuilder
                .Append("select * ")
                .Append("from ")
                .Append("[移行DB].[dbo].[MIG_" & rep.ToString & "_項目情報_" & ReportInputType.エントリー.ToString & "]")
            End With
            columnData = GetDBData(sqlBuilder.ToString)

            Dim headerNo As New DataTable
            sqlBuilder.Clear()
            With sqlBuilder
                .Append("select ID ")
                .Append("from ")
                .Append("[移行DB].[dbo].[MIG_" & rep.ToString & "_報告書_" & ReportInputType.ヘッダ.ToString & "] ")
                .Append("where ID = ")
                .Append("(select max(ID) ")
                .Append("from ")
                .Append("[移行DB].[dbo].[MIG_" & rep.ToString & "_報告書_" & ReportInputType.ヘッダ.ToString & "])")
            End With
            headerNo = GetDBData(sqlBuilder.ToString)

            Do While (Not IsNothing(xlSheet.Range("B" & currentRow).Value))
                'ページ終端行ならスキップ
                If Not xlSheet.Range("B" & currentRow).Value.ToString = "1" Then
                    Dim reportData As New Dictionary(Of String, String)
                    reportData.Add("報告書管理番号", headerNo.Rows(0)(0))
                    For Each row As DataRow In columnData.Rows
                        Dim arrData(,) As Object = xlSheet.Range(xlSheet.Cells(currentRow, row("列番号")), xlSheet.Cells(currentRow, row("列番号") + row("文字数") - 1)).Value
                        reportData.Add(row("列名"), CombineCharacter(arrData))
                    Next
                    dTable.Rows.Add(GetFormatReportData(dTable, reportData))

                End If

                '次の行の分カウント増加
                currentRow = currentRow + 2

            Loop

            Call Update(dTable)

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

    ''' <summary>
    ''' 文字列を結合します。
    ''' </summary>
    ''' <param name="target">結合対象。</param>
    ''' <returns></returns>
    Private Function CombineCharacter(ByVal target As Object(,)) As String
        Dim combinedString As String = ""
        For Each v As String In target
            combinedString = combinedString & v
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

            ' mExcelApp解放
            If Not mExcelApp Is Nothing Then
                Try
                    ' アラートを戻す
                    mExcelApp.DisplayAlerts = True
                    mExcelApp.Quit()
                Finally
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mExcelApp)
                    mExcelApp = Nothing
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

    '        ' mExcelApp解放
    '        If Not mExcelApp Is Nothing Then
    '            Try
    '                ' アラートを戻す
    '                mExcelApp.DisplayAlerts = True
    '                mExcelApp.Quit()
    '            Finally
    '                System.Runtime.InteropServices.Marshal.ReleaseComObject(mExcelApp)
    '                mExcelApp = Nothing
    '            End Try
    '        End If

    '    Catch ex As Exception
    '        Throw
    '    End Try

    'End Sub
#End Region


End Class

