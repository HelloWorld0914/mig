Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Data

Public Class DBAccess
    Public Shared gSqlConn As SqlConnection = Nothing
    Public Shared gSqlAdpt As SqlDataAdapter = Nothing

    ''' <summary>
    ''' 接続を開きます。
    ''' </summary>
    ''' <returns></returns>
    Private Function OpenConnection() As Boolean
        Try
            If Not CheckConnectionOpend() Then
                gSqlConn = New SqlConnection(ConfigurationManager.ConnectionStrings("MIG.MySettings.ConnectionString").ConnectionString)
                gSqlConn.Open()
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
                gSqlConn.Close()
                gSqlConn.Dispose()
                gSqlConn = Nothing
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
            If IsNothing(gSqlConn) Then Return False
            If gSqlConn.State <> ConnectionState.Open Then
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

            Using cmd As New SqlCommand(sql, gSqlConn)
                If clone Then
                    'メンバ変数使用
                    gSqlAdpt = New SqlDataAdapter
                    gSqlAdpt.SelectCommand = cmd
                    gSqlAdpt.Fill(dTable)
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
            Dim cmdBuilder As New SqlCommandBuilder(gSqlAdpt)
            gSqlAdpt.Update(dTable)

            gSqlAdpt.Dispose()
            gSqlAdpt = Nothing

        Catch ex As Exception
            Throw
        End Try
    End Sub

End Class
