Public Class Form3
    Dim count1 As Integer = 0
    Dim count2 As Integer = 0
    Dim count3 As Integer = 0
    Dim count4 As Integer = 0

    Public Sub DrawPieChart(ByVal percents() As Integer, ByVal colors() As Color, _
ByVal surface As Graphics, ByVal location As Point, ByVal pieSize As Size)
        Dim sum As Integer = 0
        For Each percent As Integer In percents
            sum += percent
        Next
        Dim percentTotal As Integer = 0
        For percent As Integer = 0 To percents.Length() - 1
            surface.FillPie( _
            New SolidBrush(colors(percent)), _
            New Rectangle(location, pieSize), CType(percentTotal * 360 / 100, Single), _
            CType(percents(percent) * 360 / 100, Single))
            percentTotal += percents(percent)
        Next
        Return
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim loopy As Integer

        For loopy = 0 To studcount - 1 Step 1
            If Studobjects(loopy).TPercent > 80 Then
                count1 = count1 + 1
            ElseIf Studobjects(loopy).TPercent > 75 And Studobjects(loopy).TPercent <= 80 Then
                count2 = count2 + 1
            ElseIf Studobjects(loopy).TPercent > 60 And Studobjects(loopy).TPercent <= 75 Then
                count3 = count3 + 1
            ElseIf Studobjects(loopy).TPercent <= 60 Then
                count4 = count4 + 1
            End If
        Next
        Label8.Text = count1.ToString
        Label9.Text = count2.ToString
        Label10.Text = count3.ToString
        Label11.Text = count4.ToString

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
      
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim percents() As Integer = {count1.ToString, count2.ToString, count3.ToString, count4.ToString}
        Dim colors() As Color = {Color.Blue, Color.Green, Color.Red}
        Dim graphics As Graphics = Me.CreateGraphics
        Dim location As Point = New Point(0, 0)
        Dim size As Size = New Size(200, 200)
        DrawPieChart(percents, colors, graphics, location, size)
    End Sub
End Class