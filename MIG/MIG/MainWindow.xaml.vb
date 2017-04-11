Class MainWindow
    Private Sub btn_Click(sender As Object, e As RoutedEventArgs) Handles btn.Click
        Dim o As New OperateExcel

        Try
            o.LoadReport(cmb.SelectedValue.content)

            Me.DataContext = o.GetExcelData

        Catch ex As Exception
            MessageBox.Show(ex.Message)

        Finally
            o.Dispose()
            'o = Nothing
            'Me.Close()
        End Try

    End Sub
End Class
