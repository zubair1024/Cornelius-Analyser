﻿Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim frm As New Form2
        frm.Show()

        'Me.Hide()
        'Form2.Show()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
       
    End Sub
End Class
