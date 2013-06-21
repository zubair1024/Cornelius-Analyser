
Public Class Form3
    Dim count1 As Integer = 0
    Dim Studobjects(200) As studentobject
    Dim indexx As Integer = -1
    Dim studcount As Integer = 0

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
    Dim Tper() As Integer
    Dim cc As Integer



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
        ' For k = 0 To studcount - 1 Step 1
        'ListBox1.Items.Add(Studobjects(k).sname)
        'Next
        ' For k = 0 To studcount - 1 Step 1
        'Tper(k) = Studobjects(k).TPercent
        'cc = k + 1
        'Next

        ' For k = 0 To studcount - 1 Step 1
        ' If ListBox1.SelectedIndex = k Then
        'Chart1.Series(1).Points.Add(Studobjects(k).TPercent)
        ' Chart1.Series(1).Label = "Sem 5"
        ' End If
        ' Next

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
        For k = 0 To studcount - 1 Step 1
            ListBox1.Items.Add(Studobjects(k).sname)
        Next
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

        'Chart3.Series(0).Points.Add(Label8.Text.ToString)
        'Chart3.Series(0).Points.Last.Label = ">80"
        'Chart3.Series(0).Points.Last.Color = Color.Blue
        'Chart3.Series(0).Points.Add(Label9.Text.ToString)
        'Chart3.Series(0).Points.Last.Label = "75-80"
        'Chart3.Series(0).Points.Last.Color = Color.Green
        'Chart3.Series(0).Points.Add(Label10.Text.ToString)
        'Chart3.Series(0).Points.Last.Label = "60-75"
        'Chart3.Series(0).Points.Last.Color = Color.Red
        'Chart3.Series(0).Points.Add(Label11.Text.ToString)
        'Chart3.Series(0).Points.Last.Label = "<60"
        'Chart3.Series(0).Points.Last.Color = Color.Yellow
        If Label8.Text.ToString = 0 And Label9.Text.ToString = 0 And Label10.Text.ToString = 0 And Label11.Text.ToString = 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.LightBlue
            Chart3.Series(0).Points.Add(Label9.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "75-80"
            Chart3.Series(0).Points.Last.Color = Color.Green
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "60-75"
            Chart3.Series(0).Points.Last.Color = Color.Red
            Chart3.Series(0).Points.Add(Label11.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "<60"
            Chart3.Series(0).Points.Last.Color = Color.Yellow
        ElseIf Label8.Text.ToString = 0 And Label9.Text.ToString = 0 And Label10.Text.ToString = 0 And Label11.Text.ToString <> 0 Then
            Chart3.Series(0).Points.Add(Label11.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "<60"
            Chart3.Series(0).Points.Last.Color = Color.Yellow
        ElseIf Label8.Text.ToString = 0 And Label9.Text.ToString = 0 And Label10.Text.ToString <> 0 And Label11.Text.ToString = 0 Then
           
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "60-75"
            Chart3.Series(0).Points.Last.Color = Color.Red

        ElseIf Label8.Text.ToString = 0 And Label9.Text.ToString = 0 And Label10.Text.ToString <> 0 And Label11.Text.ToString <> 0 Then
           
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "60-75"
            Chart3.Series(0).Points.Last.Color = Color.Red
            Chart3.Series(0).Points.Add(Label11.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "<60"
            Chart3.Series(0).Points.Last.Color = Color.Yellow
        ElseIf Label8.Text.ToString = 0 And Label9.Text.ToString <> 0 And Label10.Text.ToString = 0 And Label11.Text.ToString = 0 Then

            
            Chart3.Series(0).Points.Add(Label9.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "75-80"
            Chart3.Series(0).Points.Last.Color = Color.Green
          

        ElseIf Label8.Text.ToString = 0 And Label9.Text.ToString = 1 And Label10.Text.ToString = 0 And Label11.Text.ToString = 1 Then
            Chart1.Series(1).Points.Add(Label9.Text.ToString)
            Chart1.Series(1).Label = "75-80"
            Chart1.Series(3).Points.Add(Label11.Text.ToString)
            Chart1.Series(3).Label = "<60"

        ElseIf Label8.Text.ToString = 0 And Label9.Text.ToString <> 0 And Label10.Text.ToString <> 0 And Label11.Text.ToString = 0 Then
    
            Chart3.Series(0).Points.Add(Label9.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "75-80"
            Chart3.Series(0).Points.Last.Color = Color.Green
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "60-75"
            Chart3.Series(0).Points.Last.Color = Color.Red
 
        ElseIf Label8.Text.ToString = 0 And Label9.Text.ToString <> 0 And Label10.Text.ToString <> 0 And Label11.Text.ToString <> 0 Then
            Chart1.Series(1).Points.Add(Label9.Text.ToString)
            Chart1.Series(1).Label = "75-80"
            Chart1.Series(2).Points.Add(Label10.Text.ToString)
            Chart1.Series(2).Label = "60-75"
            Chart1.Series(3).Points.Add(Label11.Text.ToString)
            Chart1.Series(3).Label = "<60"
        ElseIf Label8.Text.ToString <> 0 And Label9.Text.ToString = 0 And Label10.Text.ToString = 0 And Label11.Text.ToString = 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.Blue
            

        ElseIf Label8.Text.ToString <> 0 And Label9.Text.ToString = 0 And Label10.Text.ToString = 0 And Label11.Text.ToString <> 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.Blue
            Chart3.Series(0).Points.Add(Label11.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "<60"
            Chart3.Series(0).Points.Last.Color = Color.Yellow
        ElseIf Label8.Text.ToString <> 0 And Label9.Text.ToString = 0 And Label10.Text.ToString <> 0 And Label11.Text.ToString = 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.Blue
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "60-75"
            Chart3.Series(0).Points.Last.Color = Color.Red
          
        ElseIf Label8.Text.ToString <> 0 And Label9.Text.ToString = 0 And Label10.Text.ToString <> 0 And Label11.Text.ToString <> 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.Blue
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "60-75"
            Chart3.Series(0).Points.Last.Color = Color.Red
            Chart3.Series(0).Points.Add(Label11.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "<60"
            Chart3.Series(0).Points.Last.Color = Color.Yellow
        ElseIf Label8.Text.ToString <> 0 And Label9.Text.ToString <> 0 And Label10.Text.ToString = 0 And Label11.Text.ToString = 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.Blue
            Chart3.Series(0).Points.Add(Label9.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "75-80"
            Chart3.Series(0).Points.Last.Color = Color.Green
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            '
        ElseIf Label8.Text.ToString <> 0 And Label9.Text.ToString <> 0 And Label10.Text.ToString = 0 And Label11.Text.ToString <> 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.Blue
            Chart3.Series(0).Points.Add(Label9.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "75-80"
            Chart3.Series(0).Points.Last.Color = Color.Green
            Chart3.Series(0).Points.Add(Label11.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "<60"
            Chart3.Series(0).Points.Last.Color = Color.Yellow
            '
        ElseIf Label8.Text.ToString <> 0 And Label9.Text.ToString <> 0 And Label10.Text.ToString <> 0 And Label11.Text.ToString = 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.Blue
            Chart3.Series(0).Points.Add(Label9.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "75-80"
            Chart3.Series(0).Points.Last.Color = Color.Green
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "60-75"
            Chart3.Series(0).Points.Last.Color = Color.Red

        ElseIf Label8.Text.ToString <> 0 And Label9.Text.ToString <> 0 And Label10.Text.ToString <> 0 And Label11.Text.ToString <> 0 Then
            Chart3.Series(0).Points.Add(Label8.Text.ToString)
            Chart3.Series(0).Points.Last.Label = ">80"
            Chart3.Series(0).Points.Last.Color = Color.Blue


            Chart3.Series(0).Points.Add(Label9.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "75-80"
            Chart3.Series(0).Points.Last.Color = Color.Green
            Chart3.Series(0).Points.Add(Label10.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "60-75"
            Chart3.Series(0).Points.Last.Color = Color.Red
            Chart3.Series(0).Points.Add(Label11.Text.ToString)
            Chart3.Series(0).Points.Last.Label = "<60"
            Chart3.Series(0).Points.Last.Color = Color.Yellow
        End If
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
        Chart4.Series(0).Points.Last.Color = Color.LightGreen
        Chart4.Series(0).Points.Last.Label = "PASSED"
        Chart4.Series(0).Points.Add(Label15.Text.ToString)
        Chart4.Series(0).Points.Last.Label = "FAILED"
        Chart4.Series(0).Points.Last.Color = Color.LightSalmon

    End Sub

  
    Private Sub Chart5_Click(sender As Object, e As EventArgs)

    End Sub

  
    Private Sub Chart5_Click_1(sender As Object, e As EventArgs) Handles Chart5.Click

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim PF As Integer = 0
        Dim D As Integer = 0
        Dim no As Integer = 0

        'To empty the chart
        '    For Each z As DataVisualization.Charting.ChartElementCollection(Of String) In Chart5.Series
        'z.Dispose()
        ' Next

        Me.Chart5.Series.Clear()



        indexx = ListBox1.SelectedIndex()
        If ListBox1.Items.Count = 1 Then
            indexx = 0
        End If
        Label53.ForeColor = SystemColors.ControlText
        Label52.ForeColor = SystemColors.ControlText
        Label51.ForeColor = SystemColors.ControlText
        Label50.ForeColor = SystemColors.ControlText
        Label49.ForeColor = SystemColors.ControlText
        Label48.ForeColor = SystemColors.ControlText
        Label56.ForeColor = SystemColors.ControlText
        Label55.ForeColor = SystemColors.ControlText
        Label66.ForeColor = SystemColors.ControlText

        Label31.Text = Studobjects(0).Subject(0)
        Label32.Text = Studobjects(0).Subject(1)
        Label33.Text = Studobjects(0).Subject(2)
        Label34.Text = Studobjects(0).Subject(3)
        Label35.Text = Studobjects(0).Subject(4)
        Label36.Text = Studobjects(0).Subject(5)
        Label37.Text = Studobjects(0).Subject(6)
        Label38.Text = Studobjects(0).Subject(7)

        Label53.Text = Studobjects(indexx).SubjectT(0)
        If Studobjects(indexx).SubjectT(0) < 40 Then
            PF = 1
            Label53.ForeColor = Color.Red
            Label53.Text = Studobjects(indexx).SubjectT(0) + "*"
        End If
        Label52.Text = Studobjects(indexx).SubjectT(1)
        If Studobjects(indexx).SubjectT(1) < 40 Then
            PF = 1
            Label52.ForeColor = Color.Red
            Label52.Text = Studobjects(indexx).SubjectT(1) + "*"
        End If
        Label51.Text = Studobjects(indexx).SubjectT(2)
        If Studobjects(indexx).SubjectT(2) < 40 Then
            PF = 1
            Label51.ForeColor = Color.Red
            Label51.Text = Studobjects(indexx).SubjectT(2) + "*"
        End If
        Label50.Text = Studobjects(indexx).SubjectT(3)
        If Studobjects(indexx).SubjectT(3) < 40 Then
            PF = 1
            Label50.ForeColor = Color.Red
            Label50.Text = Studobjects(indexx).SubjectT(3) + "*"
        End If
        Label49.Text = Studobjects(indexx).SubjectT(4)
        If Studobjects(indexx).SubjectT(4) < 40 Then
            PF = 1
            Label49.ForeColor = Color.Red
            Label49.Text = Studobjects(indexx).SubjectT(4) + "*"
        End If
        Label48.Text = Studobjects(indexx).SubjectT(5)
        If Studobjects(indexx).SubjectT(5) < 40 Then
            PF = 1
            Label48.ForeColor = Color.Red
            Label48.Text = Studobjects(indexx).SubjectT(5) + "*"
        End If
        Label56.Text = Studobjects(indexx).SubjectT(6)
        If Studobjects(indexx).SubjectT(6) < 40 Then
            PF = 1
            Label56.ForeColor = Color.Red
            Label56.Text = Studobjects(indexx).SubjectT(6) + "*"
        End If
        Label55.Text = Studobjects(indexx).SubjectT(7)
        If Studobjects(indexx).SubjectT(7) < 40 Then
            PF = 1
            Label55.ForeColor = Color.Red
            Label55.Text = Studobjects(indexx).SubjectT(7) + "*"
        End If

        'For no = 0 To Studobjects(indexx).Subject.Length
        ' If Studobjects(indexx).SubjectT(no) < 40 Then
        'PF = 1
        ' End If
        'Next
        Label39.Text = "TOTAL"
        Label67.Text = "RESULT"
        If PF = 1 Then
            Label66.ForeColor = Color.Red
            Label66.Text = Studobjects(indexx).Result
            'Label54.ForeColor = Color.Red
            Label54.Text = Studobjects(indexx).TPercent + "%"
        Else
            Label66.ForeColor = Color.Green
            Label66.Text = Studobjects(indexx).Result
            'Label54.ForeColor = Color.Green
            Label54.Text = Studobjects(indexx).TPercent + "%"
        End If

        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label53.Text.ToString)
            Chart5.Series(0).Points.Last.Label = "FAILED"
            Chart5.Series(0).Points.Last.Color = Color.LightGray

        End If
        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label14.Text.ToString)
            Chart5.Series(0).Points.Last.Label = Studobjects(0).Subject(0)
            Chart5.Series(0).Points.Last.Color = Color.LightBlue
        End If
        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label14.Text.ToString)
            Chart5.Series(0).Points.Last.Label = Studobjects(0).Subject(1)
            Chart5.Series(0).Points.Last.Color = Color.LightGoldenrodYellow

        End If
        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label14.Text.ToString)
            Chart5.Series(0).Points.Last.Label = Studobjects(0).Subject(2)
            Chart5.Series(0).Points.Last.Color = Color.LightGreen

        End If
        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label14.Text.ToString)
            Chart5.Series(0).Points.Last.Label = Studobjects(0).Subject(3)
            Chart5.Series(0).Points.Last.Color = Color.LightPink
        End If
        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label14.Text.ToString)
            Chart5.Series(0).Points.Last.Label = Studobjects(0).Subject(4)
            Chart5.Series(0).Points.Last.Color = Color.LightSalmon
        End If
        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label14.Text.ToString)
            Chart5.Series(0).Points.Last.Label = Studobjects(0).Subject(5)
            Chart5.Series(0).Points.Last.Color = Color.LightYellow

        End If
        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label14.Text.ToString)
            Chart5.Series(0).Points.Last.Label = Studobjects(0).Subject(6)
            Chart5.Series(0).Points.Last.Color = Color.LimeGreen

        End If
        If Label53.Text.ToString > 0 Then
            Chart5.Series(0).Points.Add(Label14.Text.ToString)
            Chart5.Series(0).Points.Last.Label = Studobjects(0).Subject(7)
            Chart5.Series(0).Points.Last.Color = Color.LightSteelBlue

        End If

    End Sub
End Class