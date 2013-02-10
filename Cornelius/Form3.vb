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
    Dim passpercent As Single = 0
    Dim failpercent As Single = 0
    Dim Passlocal, Faillocal As Integer


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
        'For loopy = 0 To studcount - 1 Step 1
        'If Studobjects(loopy).Result = "Failed" Then
        ' count1 = count1 + 1

        'End If
        'Next
        Label8.Text = count1.ToString
        Label9.Text = count2.ToString
        Label10.Text = count3.ToString
        Label11.Text = count4.ToString

        'passpercent = 100 * (Pass / totall)
        'passpercent = 100 * (Fail / totall)
        Passlocal = Form2.Pass
        Faillocal = Form2.Fail
        Label14.Text = Passlocal
        Label15.Text = Faillocal

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
        ''Dim percents() As Integer = {count1percent.ToString, count2percent.ToString, count3percent.ToString, count4percent.ToString}
        '' Dim colors() As Color = {Color.Blue, Color.Green, Color.Red, Color.Yellow}
        'Dim graphics As Graphics = Me.CreateGraphics
        'Dim location As Point = New Point(350, 20)
        ' Dim size As Size = New Size(200, 200)
        'DrawPieChart(percents, colors, graphics, location, size)
        'Label1.Text = "Pie Chart"
        'Label23.Text = Label8.Text
        'Label24.Text = Label9.Text
        'Label25.Text = Label10.Text
        'Label26.Text = Label11.Text
        Chart1.Series(0).Points.Add(Label8.Text.ToString)
        Chart1.Series(0).Label = ">80"
        Chart1.Series(1).Points.Add(Label9.Text.ToString)
        Chart1.Series(1).Label = "75-80"
        Chart1.Series(2).Points.Add(Label10.Text.ToString)
        Chart1.Series(2).Label = "60-75"
        Chart1.Series(3).Points.Add(Label11.Text.ToString)
        Chart1.Series(3).Label = "<60"

        Chart3.Series(0).Points.Add(Label8.Text.ToString)
        Chart3.Series(0).Points.Last.Label = ">80"
        Chart3.Series(0).Points.Add(Label9.Text.ToString)
        Chart3.Series(0).Points.Last.Label = "75-80"
        Chart3.Series(0).Points.Add(Label10.Text.ToString)
        Chart3.Series(0).Points.Last.Label = "60-75"
        Chart3.Series(0).Points.Add(Label11.Text.ToString)
        Chart3.Series(0).Points.Last.Label = "<60"
        Button1.Enabled = False
    End Sub






    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Passlocal = Form2.Pass
        Faillocal = Form2.Fail
        totall = Passlocal + Faillocal
        passpercent = 100 * (Passlocal / totall)
        failpercent = 100 * (Faillocal / totall)
        ' Dim percents() As Integer = {passpercent.ToString, failpercent.ToString}
        'Dim colors() As Color = {Color.Green, Color.Red}
        'Dim graphics As Graphics = Me.CreateGraphics
        'Dim location As Point = New Point(350, 248)
        'Dim size As Size = New Size(200, 200)
        'DrawPieChart(percents, colors, graphics, location, size)
        'Label1.Text = "Pie Chart"
        Label28.Text = Label14.Text
        Label27.Text = Label15.Text
        Chart2.Series(0).Points.Add(Label14.Text.ToString)
        Chart2.Series(0).Label = "PASSED"
        Chart2.Series(1).Points.Add(Label15.Text.ToString)
        Chart2.Series(1).Label = "FAILED"
        Button2.Enabled = False

        Chart4.Series(0).Points.Add(Label14.Text.ToString)
        Chart4.Series(0).Points.Last.Label = "PASSED"
        Chart4.Series(0).Points.Add(Label15.Text.ToString)
        Chart4.Series(0).Points.Last.Label = "FAILED"
    End Sub
End Class