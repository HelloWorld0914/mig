Imports Microsoft.WindowsAPICodePack.Dialogs

Class MainWindow
    Private Sub GetFolderName(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim dlg = New CommonOpenFileDialog("保存フォルダ選択")
        dlg.IsFolderPicker = True
        Dim ret = dlg.ShowDialog()
        If ret = CommonFileDialogResult.Ok Then
            Me.txt_FolderName.Text = dlg.FileName
        End If

        'Dim o As New OperateExcel_NetOffice
        'o.MigrateReportData("C:\Users\nesi\Desktop\ICR(178) RD(KSKSより）.xls")

        Dim o As New OperateExcel_NetOffice
        Call o.MigrateReportData(dlg.FileName)

    End Sub

    'Private Sub btn_Click(sender As Object, e As RoutedEventArgs) Handles btn.Click
    '    Dim o As New OperateExcel

    '    Try
    '        o.LoadReport(cmb.SelectedValue.content)


    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)

    '    Finally
    '        o.Dispose()
    '        'o = Nothing
    '        'Me.Close()
    '    End Try

    'End Sub
End Class
