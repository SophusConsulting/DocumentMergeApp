Module MainArgModule

    Sub main(ByVal args() As String)
        If args.Length > 0 Then
            Dim frmmain As New MainForm()
            frmmain.Show()
        End If
    End Sub

End Module
