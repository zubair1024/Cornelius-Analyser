Public Class Form3
    Dim count1 As Integer = 0
    Dim count2 As Integer = 0
    Dim count3 As Integer = 0
    Dim count4 As Integer = 0
    Dim totall As Integer = 0
    Dim count1percent As Single = 0
    Dim count2percent As Single = 0
    Dim count3percent As Single = 0
    Dim count4percent As Single = 0

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
    Public Structure studentobject
        Dim rollno As String
        Dim sname As String
        Dim temp As String
        Dim Subject() As String
        Dim SubjectC() As String
        Dim SubjectA() As String
        Dim SubjectE() As String
        Dim SubjectT() As String
        Dim Aggregate As String
        Dim TPercent As String
        Dim Result As String
        Dim scount As Integer

    End Structure
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal Studobjects() As Form2.studentobject, ByVal studcount As Integer)
        InitializeComponent()

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

        totall = count1 + count2 + count3 + count4
        count1percent = 100 * (count1 / totall)
        count2percent = 100 * (count2 / totall)
        count3percent = 100 * (count3 / totall)
        count4percent = 100 * (count4 / totall)

    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim percents() As Integer = {count1percent.ToString, count2percent.ToString, count3percent.ToString, count4percent.ToString}
        Dim colors() As Color = {Color.Blue, Color.Green, Color.Red, Color.Yellow}
        Dim graphics As Graphics = Me.CreateGraphics
        Dim location As Point = New Point(350, 50)
        Dim size As Size = New Size(200, 200)
        DrawPieChart(percents, colors, graphics, location, size)
        Label1.Text = "Pie Chart"
    End Sub
End Class