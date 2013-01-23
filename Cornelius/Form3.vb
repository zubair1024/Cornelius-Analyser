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


    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = True Then
            totall = count1 + count2 + count3
            count1percent = 100 * (count1 / totall)
            count2percent = 100 * (count2 / totall)
            count3percent = 100 * (count3 / totall)
            Dim percents() As Integer = {count1percent.ToString, count2percent.ToString, count3percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Green, Color.Red}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = True And CheckBox3.Checked = True Then
            totall = count2 + count3
            count2percent = 100 * (count2 / totall)
            count3percent = 100 * (count3 / totall)
            Dim percents() As Integer = {count2percent.ToString, count3percent.ToString}
            Dim colors() As Color = {Color.Green, Color.Red}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox3.Checked = True Then
            totall = count1 + count3
            count1percent = 100 * (count1 / totall)
            count3percent = 100 * (count3 / totall)
            Dim percents() As Integer = {count1percent.ToString, count3percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Red}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = False Then
            totall = count1 + count2
            count1percent = 100 * (count1 / totall)
            count2percent = 100 * (count2 / totall)
            Dim percents() As Integer = {count1percent.ToString, count2percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Green}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = True Then
            totall = count3
            count3percent = 100 * (count3 / totall)
            Dim percents() As Integer = {count3percent.ToString}
            Dim colors() As Color = {Color.Red}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = True And CheckBox3.Checked = False Then
            totall = count2
            count2percent = 100 * (count2 / totall)
            Dim percents() As Integer = {count2percent.ToString}
            Dim colors() As Color = {Color.Green}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox3.Checked = False Then
            totall = count1
            count1percent = 100 * (count1 / totall)
            Dim percents() As Integer = {count1percent.ToString}
            Dim colors() As Color = {Color.Blue}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = True And CheckBox1.Checked = False And CheckBox2.Checked = True And CheckBox3.Checked = True Then
            totall = count2 + count3 + count4
            count2percent = 100 * (count2 / totall)
            count3percent = 100 * (count3 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count2percent.ToString, count3percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Green, Color.Red, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = True And CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox3.Checked = True Then
            totall = count1 + count3 + count4
            count1percent = 100 * (count1 / totall)
            count3percent = 100 * (count3 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count1percent.ToString, count3percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Red, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = True And CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox3.Checked = False Then
            totall = count1 + count2
            count1percent = 100 * (count1 / totall)
            count2percent = 100 * (count2 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count1percent.ToString, count2percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Green, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = True And CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = True Then
            totall = count3 + count4
            count3percent = 100 * (count3 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count3percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Red, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = True And CheckBox1.Checked = False And CheckBox2.Checked = True And CheckBox3.Checked = False Then
            totall = count2 + count4
            count2percent = 100 * (count2 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count2percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Green, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        ElseIf CheckBox4.Checked = True And CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox3.Checked = False Then
            totall = count1 + count4
            count1percent = 100 * (count1 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count1percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox4.Checked = True Then
            totall = count1 + count2 + count4
            count1percent = 100 * (count1 / totall)
            count2percent = 100 * (count2 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count1percent.ToString, count2percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Green, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox3.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = True And CheckBox4.Checked = True Then
            totall = count2 + count4
            count2percent = 100 * (count2 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count2percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Green, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox3.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox4.Checked = True Then
            totall = count1 + count4
            count1percent = 100 * (count1 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count1percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox3.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = True And CheckBox4.Checked = False Then
            totall = count1 + count2
            count1percent = 100 * (count1 / totall)
            count2percent = 100 * (count2 / totall)
            Dim percents() As Integer = {count1percent.ToString, count2percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Green}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox3.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox4.Checked = True Then
            totall = count4
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count4percent.ToString}
            Dim colors() As Color = {Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox3.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = True And CheckBox4.Checked = False Then
            totall = count2
            count2percent = 100 * (count2 / totall)
            Dim percents() As Integer = {count2percent.ToString}
            Dim colors() As Color = {Color.Green}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox3.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox4.Checked = False Then
            totall = count1
            count1percent = 100 * (count1 / totall)
            Dim percents() As Integer = {count1percent.ToString}
            Dim colors() As Color = {Color.Blue}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = False And CheckBox1.Checked = True And CheckBox4.Checked = True And CheckBox3.Checked = True Then
            totall = count1 + count2 + count3
            count1percent = 100 * (count1 / totall)
            count4percent = 100 * (count4 / totall)
            count3percent = 100 * (count3 / totall)
            Dim percents() As Integer = {count1percent.ToString, count4percent.ToString, count3percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Yellow, Color.Red}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox2.Checked = False And CheckBox1.Checked = False And CheckBox4.Checked = True And CheckBox3.Checked = True Then
            totall = count4 + count3
            count4percent = 100 * (count4 / totall)
            count3percent = 100 * (count3 / totall)
            Dim percents() As Integer = {count4percent.ToString, count3percent.ToString}
            Dim colors() As Color = {Color.Yellow, Color.Red}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox2.Checked = False And CheckBox1.Checked = True And CheckBox4.Checked = False And CheckBox3.Checked = True Then
            totall = count1 + count3
            count1percent = 100 * (count1 / totall)
            count3percent = 100 * (count3 / totall)
            Dim percents() As Integer = {count1percent.ToString, count3percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Red}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox2.Checked = False And CheckBox1.Checked = True And CheckBox4.Checked = True And CheckBox3.Checked = False Then
            totall = count1 + count4
            count1percent = 100 * (count1 / totall)
            count4percent = 100 * (count4 / totall)
            Dim percents() As Integer = {count1percent.ToString, count4percent.ToString}
            Dim colors() As Color = {Color.Blue, Color.Yellow}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox4.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = True Then
            totall = count3
            count3percent = 100 * (count3 / totall)
            Dim percents() As Integer = {count3percent.ToString}
            Dim colors() As Color = {Color.Red}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox4.Checked = False And CheckBox1.Checked = False And CheckBox2.Checked = True And CheckBox3.Checked = False Then
            totall = count2
            count2percent = 100 * (count2 / totall)
            Dim percents() As Integer = {count2percent.ToString}
            Dim colors() As Color = {Color.Green}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
        If CheckBox4.Checked = False And CheckBox1.Checked = True And CheckBox2.Checked = False And CheckBox3.Checked = False Then
            totall = count1
            count1percent = 100 * (count1 / totall)
            Dim percents() As Integer = {count1percent.ToString}
            Dim colors() As Color = {Color.Blue}
            Dim graphics As Graphics = Me.CreateGraphics
            Dim location As Point = New Point(350, 50)
            Dim size As Size = New Size(200, 200)
            DrawPieChart(percents, colors, graphics, location, size)
            Label1.Text = "Pie Chart"
        End If
    End Sub
End Class