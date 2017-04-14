Imports Microsoft.WindowsAPICodePack.Dialogs

Class MainWindow
    Private Sub GetFolderName(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim dlg = New CommonOpenFileDialog("保存フォルダ選択")
        dlg.IsFolderPicker = True
        Dim ret = dlg.ShowDialog()
        If ret = CommonFileDialogResult.Ok Then
            Me.txt_FolderName.Text = dlg.FileName
        End If

    End Sub

    Private Sub btn_Start_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Try
            Dim o As New OperateExcel_NetOffice
            Call o.MigrateReportData(txt_FolderName.Text)

        Catch ex As Exception
            MessageBox.Show("フォルダ名を入力してください。")
        End Try
    End Sub

End Class
