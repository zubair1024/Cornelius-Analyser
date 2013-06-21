'Imports Microsoft.Office.Interop.Excel

Imports iTextSharp.text.pdf
Imports System.IO

Public Class Form2
    Dim headers As String
    Dim Studobjects(200) As studentobject
    Dim dept As String
    Public Pass As Integer = 0
    Public Fail As Integer = 0
   
    Dim studcount As Integer = 0
    Dim minsub() As Integer
    Dim maxsub() As Integer
    Dim avgsub() As Integer
    Dim sumsub() As Integer
    Dim indexx As Integer = -1

    Private Sub Browsebtn_Click(sender As Object, e As EventArgs) Handles Browsebtn.Click
        Button3.Enabled = True

        If (OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK) Then

            Dim sourcePdf As String = OpenFileDialog1.FileName
            Dim traf = New iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf)
            Dim treader = New iTextSharp.text.pdf.PdfReader(traf, Nothing)
            Dim tpageCount = treader.NumberOfPages
            Dim i As Integer = 1
            Dim data As String = ""
            Dim tempdata As String
            tempdata = ReadPdfFile(OpenFileDialog1.FileName)
            'For debug only 
            '   Dim fwrite As New StreamWriter("C:\Web\test.txt")
            '  fwrite.Write(tempdata)
            ' fwrite.Close()
            ParseToObjects(tempdata)
            TextBox1.Text = sourcePdf.ToString
        End If

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
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        Button1.Enabled = False
    End Sub
    Public Function countsubjects(ByVal rawdata As String)
        Dim subcount As Integer = 0
        rawdata = rawdata.Substring(0, rawdata.IndexOf("Register"))
        Dim i As Integer = rawdata.IndexOf("Subject")
        i = i + 1
        While i > 0
            i = rawdata.IndexOf("Subject", i + 1)
            subcount = subcount + 1
        End While
        Return subcount
    End Function
    Public Sub ParseToObjects(ByVal rawdata As String)
        Dim i As Integer

        i = rawdata.IndexOf("Year")
        headers = rawdata.Substring(0, i + 4 + 6)
        rawdata = rawdata.Substring(i + 10)
        rawdata = rawdata.TrimStart()
        i = headers.IndexOf("BT")
        If i = -1 Then
            i = headers.IndexOf("MT")
        End If
        dept = headers.Substring(i + 2, 2)
        i = rawdata.IndexOf(dept)
        Dim j As Integer = 0
        Dim k As Integer = 0
         Dim l As Integer = rawdata.IndexOf("Total Mark")
        While (l <> -1)
            k = rawdata.IndexOf("CIA")
            If k = -1 Then
                GoTo skip
            End If
            k = rawdata.IndexOf("CIA", k + 4)
            k = rawdata.IndexOf(dept, k + 3)
            If i <> 0 Then
                Studobjects(j).temp = rawdata.Substring(i - 9, (k - 9) - (i - 9))
            Else
                Studobjects(j).temp = rawdata.Substring(0, k - 9)

            End If
            rawdata = rawdata.Substring(k - 9)
            j = j + 1
            i = rawdata.IndexOf(dept)
            l = rawdata.IndexOf("Total Mark")
        End While

skip:   Label10.Text = j.ToString
        studcount = j
        CleanData(j)
    End Sub
    Public Sub CleanData(ByVal count As Integer)
        Dim k As Integer = 0
        Dim parray(10) As Integer
        Dim farray(10) As Integer

        For i = 0 To count - 1 Step 1
            k = Studobjects(i).temp.IndexOf("Card")
            If k <> -1 Then
                Studobjects(i).temp = Studobjects(i).temp.Substring(k + 4, Studobjects(i).temp.Length - (k + 4))
                Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            End If
            k = 0
            Studobjects(i).Aggregate = Studobjects(i).temp.Substring(0, 3)
            Studobjects(i).temp = Studobjects(i).temp.Substring(7, Studobjects(i).temp.Length - 7)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            Studobjects(i).scount = 0
step1:      Dim j As Integer = Studobjects(i).temp.IndexOf(dept)
            If j <> -1 Then
                Studobjects(i).scount = Studobjects(i).scount + 1
                j = Studobjects(i).temp.IndexOf(dept, 5)
                If j <> -1 Then
                    ReDim Preserve Studobjects(i).Subject(k)
                    Studobjects(i).Subject(k) = New String(Studobjects(i).temp.Substring(0, j))
                    Studobjects(i).temp = Studobjects(i).temp.Substring(j, Studobjects(i).temp.Length - j)
                Else
                    ReDim Preserve Studobjects(i).Subject(k)
                    Studobjects(i).Subject(k) = New String(Studobjects(i).temp.Substring(0, 8))
                    Studobjects(i).temp = Studobjects(i).temp.Substring(8, Studobjects(i).temp.Length - 8)

                End If
                k = k + 1
                GoTo step1
            End If
            j = Studobjects(i).temp.IndexOf("Distinction")

            If j = -1 Then
                j = Studobjects(i).temp.IndexOf("First Class")
                If j = -1 Then
                    j = Studobjects(i).temp.IndexOf("Pass Class")
                    If j = -1 Then
                        j = Studobjects(i).temp.IndexOf("FAILED")
                        Studobjects(i).Result = "FAILED"
                        k = Studobjects(i).temp.IndexOf("MAX")
                        Fail = Fail + 1

                        If k > j Then
                            '  Studobjects(i).TPercent = Studobjects(i).temp.Substring(0, j)

                        Else
                            'Studobjects(i).TPercent = Studobjects(i).temp.Substring(k + 3, j - (k + 3))

                        End If
                        Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("AILED") + 5)
                    Else
                        Studobjects(i).Result = "Pass Class"
                        k = Studobjects(i).temp.IndexOf("MAX")
                        'Studobjects(i).TPercent = Studobjects(i).temp.Substring(k + 3, j - (k + 3))
                        Pass = Pass + 1
                        Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Class") + 5)
                    End If
                Else
                    Studobjects(i).Result = "First Class"
                    k = Studobjects(i).temp.IndexOf("MAX")
                    'Studobjects(i).TPercent = Studobjects(i).temp.Substring(k + 3, j - (k + 3))
                    Pass = Pass + 1
                    Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Class") + 5)
                End If
            Else
                Studobjects(i).Result = "Distinction"
                k = Studobjects(i).temp.IndexOf("MAX")
                'Studobjects(i).TPercent = Studobjects(i).temp.Substring(k + 3, j - (k + 3))
                Pass = Pass + 1
                Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("ction") + 5)

            End If
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            Dim l As Integer = 0
            For k = 0 To (Studobjects(i).scount * 2) - 1 Step 1
                l = Studobjects(i).temp.IndexOf(" ", l + 1)

            Next
            Studobjects(i).temp = Studobjects(i).temp.Substring(l)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart

            'Studobjects(i).temp = Studobjects(i).temp.Substring((3 * (2 * Studobjects(i).scount)) + 3)
            Studobjects(i).rollno = Studobjects(i).temp.Substring(0, 7)
            Studobjects(i).temp = Studobjects(i).temp.Substring(7)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            Studobjects(i).sname = Studobjects(i).temp.Substring(0, Studobjects(i).temp.IndexOf("ESE"))
            Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("CIA") + 9)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()
            ReDim Preserve Studobjects(i).SubjectC(Studobjects(i).scount)
            k = 0
            For k = 0 To Studobjects(i).scount - 1 Step 1
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).SubjectC(k) = Studobjects(i).temp.Substring(0, j)
                Studobjects(i).temp = Studobjects(i).temp.Substring(j)
                Studobjects(i).temp = Studobjects(i).temp.TrimStart()
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).temp = Studobjects(i).temp.Substring(j + 1)



            Next
            Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Marks") + 5)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()

            ReDim Preserve Studobjects(i).SubjectA(Studobjects(i).scount)

            For k = 0 To Studobjects(i).scount - 1 Step 1
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).SubjectA(k) = Studobjects(i).temp.Substring(0, j)
                Studobjects(i).temp = Studobjects(i).temp.Substring(j)
                Studobjects(i).temp = Studobjects(i).temp.TrimStart()
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).temp = Studobjects(i).temp.Substring(j + 1)



            Next
            Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Marks") + 5)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()

            ReDim Preserve Studobjects(i).SubjectE(Studobjects(i).scount)

            For k = 0 To Studobjects(i).scount - 1 Step 1
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).SubjectE(k) = Studobjects(i).temp.Substring(0, j)
                Studobjects(i).temp = Studobjects(i).temp.Substring(j)
                Studobjects(i).temp = Studobjects(i).temp.TrimStart()
                j = Studobjects(i).temp.IndexOf(" ")
                Studobjects(i).temp = Studobjects(i).temp.Substring(j + 1)



            Next
            Studobjects(i).temp = Studobjects(i).temp.Substring(Studobjects(i).temp.IndexOf("Mark") + 4)
            Studobjects(i).temp = Studobjects(i).temp.TrimStart()

            ReDim Preserve Studobjects(i).SubjectT(Studobjects(i).scount)
            For k = 0 To Studobjects(i).scount - 1 Step 1
                If Studobjects(i).SubjectA(k) = "" Then
                    Studobjects(i).SubjectA(k) = "0"
                End If
            Next
            For k = 0 To Studobjects(i).scount - 1 Step 1
                Dim temp(3) As Integer
                If Studobjects(i).SubjectA(k) = "AA" Then
                    temp(0) = 0
                Else
                    temp(0) = Integer.Parse(Studobjects(i).SubjectA(k))
                End If
                If Studobjects(i).SubjectC(k) = "AA" Then
                    temp(1) = 0
                Else
                    temp(1) = Integer.Parse(Studobjects(i).SubjectC(k))
                End If
                If Studobjects(i).SubjectE(k) = "AA" Then
                    temp(2) = 0
                Else
                    temp(2) = Integer.Parse(Studobjects(i).SubjectE(k))
                End If

                Studobjects(i).SubjectT(k) = temp(0) + temp(1) + temp(2)
                'j = Studobjects(i).temp.IndexOf(" ")
                'Studobjects(i).SubjectT(k) = Studobjects(i).temp.Substring(0, j)
                'Studobjects(i).temp = Studobjects(i).temp.Substring(j)
                'Studobjects(i).temp = Studobjects(i).temp.TrimStart()
                'j = Studobjects(i).temp.IndexOf(" ")
                'Studobjects(i).temp = Studobjects(i).temp.Substring(j + 1)
                Studobjects(i).TPercent = (Studobjects(i).Aggregate / (Studobjects(i).scount * 100)) * 100



            Next
            

        Next
        Label11.Text = Pass
        Label12.Text = Fail
        Label13.Text = ((Pass / (Integer.Parse(Label10.Text))) * 100).ToString("G3")
        Label14.Text = ((Fail / (Integer.Parse(Label10.Text))) * 100).ToString("G3")
        For k = 0 To studcount - 1 Step 1
            ListBox1.Items.Add(Studobjects(k).sname)
        Next
        Dim minsub() As String
        ReDim Preserve minsub(Studobjects(0).Subject.Length)
        ReDim Preserve maxsub(Studobjects(0).Subject.Length)
        ReDim Preserve sumsub(Studobjects(0).Subject.Length)
        ReDim Preserve avgsub(Studobjects(0).Subject.Length)
        For k = 0 To Studobjects(0).Subject.Length Step 1
            minsub(k) = Studobjects(0).SubjectT(k)
        Next

        Dim z As Integer = 0
        For z = 0 To Studobjects(0).Subject.Length Step 1
            For k = 0 To studcount - 1 Step 1
                If minsub(z) > Studobjects(k).SubjectT(z) Then
                    minsub(z) = Studobjects(k).SubjectT(z)
                End If
            Next
        Next
        For z = 0 To Studobjects(0).Subject.Length Step 1
            For k = 0 To studcount - 1 Step 1
                If maxsub(z) < Studobjects(k).SubjectT(z) Then
                    maxsub(z) = Studobjects(k).SubjectT(z)
                End If
            Next
        Next

        'For z = 0 To Studobjects(0).Subject.Length Step 1
        'sumsub(z) = 0
        ' Next

        'For k = 0 To studcount - 1 Step 1
        For z = 0 To Studobjects(0).Subject.Length Step 1
            For k = 0 To studcount - 1 Step 1
                sumsub(z) = sumsub(z) + Studobjects(k).SubjectT(z)
            Next
            avgsub(z) = (sumsub(z) / studcount)
        Next

        For z = 0 To Studobjects(0).Subject.Length Step 1
            For k = 0 To studcount - 1 Step 1
                If Studobjects(k).SubjectT(z) < 40 Then
                    farray(z) = farray(z) + 1

                End If
            Next
        Next

        For z = 0 To Studobjects(0).Subject.Length Step 1
            For k = 0 To studcount - 1 Step 1
                If Studobjects(k).SubjectT(z) >= 40 Then
                    parray(z) = parray(z) + 1

                End If
            Next
        Next

        Label40.Text = Studobjects(0).Subject(0)
        Label41.Text = Studobjects(0).Subject(1)
        Label42.Text = Studobjects(0).Subject(2)
        Label43.Text = Studobjects(0).Subject(3)
        Label44.Text = Studobjects(0).Subject(4)
        Label45.Text = Studobjects(0).Subject(5)
        Label46.Text = Studobjects(0).Subject(6)
        Label47.Text = Studobjects(0).Subject(7)

        Label31.Text = Studobjects(0).Subject(0)
        Label32.Text = Studobjects(0).Subject(1)
        Label33.Text = Studobjects(0).Subject(2)
        Label34.Text = Studobjects(0).Subject(3)
        Label35.Text = Studobjects(0).Subject(4)
        Label36.Text = Studobjects(0).Subject(5)
        Label37.Text = Studobjects(0).Subject(6)
        Label38.Text = Studobjects(0).Subject(7)

        Label15.Text = minsub(0)
        Label16.Text = minsub(1)
        Label17.Text = minsub(2)
        Label18.Text = minsub(3)
        Label19.Text = minsub(4)
        Label20.Text = minsub(5)
        Label21.Text = minsub(6)
        Label22.Text = minsub(7)

        Label23.Text = maxsub(0)
        Label24.Text = maxsub(1)
        Label25.Text = maxsub(2)
        Label26.Text = maxsub(3)
        Label27.Text = maxsub(4)
        Label28.Text = maxsub(5)
        Label29.Text = maxsub(6)
        Label30.Text = maxsub(7)

        Label64.Text = avgsub(0)
        Label63.Text = avgsub(1)
        Label62.Text = avgsub(2)
        Label61.Text = avgsub(3)
        Label60.Text = avgsub(4)
        Label59.Text = avgsub(5)
        Label58.Text = avgsub(6)
        Label57.Text = avgsub(7)

        Label84.Text = farray(0)
        Label83.Text = farray(1)
        Label82.Text = farray(2)
        Label81.Text = farray(3)
        Label80.Text = farray(4)
        Label79.Text = farray(5)
        Label78.Text = farray(6)
        Label77.Text = farray(7)

        Label75.Text = parray(0)
        Label74.Text = parray(1)
        Label73.Text = parray(2)
        Label72.Text = parray(3)
        Label71.Text = parray(4)
        Label70.Text = parray(5)
        Label69.Text = parray(6)
        Label68.Text = parray(7)




        Browsebtn.Enabled = False
        Button2.Enabled = True
        Button1.Enabled = True
    End Sub
    Public Function ReadPdfFile(ByVal fileName As String)

        Dim text As String = ""

        If File.Exists(fileName) Then

            Dim pdfReader As New PdfReader(fileName)

            For page As Integer = 1 To pdfReader.NumberOfPages Step 1
                ' Dim its As iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                Dim currentText As String = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(pdfReader, page)

                text = text + currentText
            Next
            pdfReader.Close()
        End If
        Return text.ToString()
    End Function

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim PF As Integer = 0
        Dim D As Integer = 0
        Dim no As Integer = 0
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

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label53.Text = "-"
        Label52.Text = "-"
        Label51.Text = "-"
        Label50.Text = "-"
        Label49.Text = "-"
        Label48.Text = "-"
        Label56.Text = "-"
        Label55.Text = "-"
        Label40.Text = "-"
        Label41.Text = "-"
        Label42.Text = "-"
        Label43.Text = "-"
        Label44.Text = "-"
        Label45.Text = "-"
        Label46.Text = "-"
        Label47.Text = "-"


        Label31.Text = "-"
        Label32.Text = "-"
        Label33.Text = "-"
        Label34.Text = "-"
        Label35.Text = "-"
        Label36.Text = "-"
        Label37.Text = "-"
        Label38.Text = "-"

        Label15.Text = "-"
        Label16.Text = "-"
        Label17.Text = "-"
        Label18.Text = "-"
        Label19.Text = "-"
        Label20.Text = "-"
        Label21.Text = "-"
        Label22.Text = "-"

        Label64.Text = "-"
        Label63.Text = "-"
        Label62.Text = "-"
        Label61.Text = "-"
        Label60.Text = "-"
        Label59.Text = "-"
        Label58.Text = "-"
        Label57.Text = "-"

        Label66.Text = "-"
        Label67.Text = "-"
        Label54.Text = "-"
        Label39.Text = "-"

        Label23.Text = "-"
        Label24.Text = "-"
        Label25.Text = "-"
        Label26.Text = "-"
        Label27.Text = "-"
        Label28.Text = "-"
        Label29.Text = "-"
        Label30.Text = "-"

        Label10.Text = "-"
        Label11.Text = "-"
        Label12.Text = "-"
        Label13.Text = "-"
        Label14.Text = "-"

        TextBox1.Text = ""

        ListBox1.Items.Clear()

        Browsebtn.Enabled = Enabled


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim frm As New Form3(Studobjects, studcount)

        frm.Show()


    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        export("Excel")
    End Sub
    Public Sub Export(ByVal format As String)

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

            Dim ExApp As Microsoft.Office.Interop.Excel.Application

            Dim ExWBook As Microsoft.Office.Interop.Excel.Workbook

            Dim ExWSheet As Microsoft.Office.Interop.Excel.Worksheet

            Dim ExRange As Microsoft.Office.Interop.Excel.Range

            Try

                ExApp = New Microsoft.Office.Interop.Excel.Application

                ExWBook = ExApp.Workbooks.Add

                ExWSheet = ExWBook.Sheets.Item(1)

            Catch ex As Exception

                MessageBox.Show("Failed to open Excel:" & ex.Message, "Export to Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try

            ExApp.DisplayAlerts = False

            ExApp.Visible = False

            'ExRange = ExWSheet.Range("A1:A" + (studcount + 1).ToString)

            ExRange = ExWSheet.Range("A1:K11")

            Dim headarray(studcount + 1, 8 + (Studobjects(0).scount * 3)) As String

            ExWSheet.Cells(1, 1) = "S.No."

            For i As Integer = 2 To studcount + 1 Step 1

                'headarray(i) = i.ToString

                'ExWSheet.Cells(i, 1) = (i - 1).ToString

            Next

            ' MessageBox.Show("B2:" + Chr(70 + (Studobjects(0).scount * 3)) + (studcount + 1).ToString)

            Dim tempval As Integer = 70 + (Studobjects(0).scount * 3)

            Dim char1, char2 As Char

            If tempval > 90 Then

                char1 = "A"

                char2 = Chr(tempval - 25)

            Else

                char1 = Chr(tempval)

                char2 = ""

            End If

            ' ExRange = ExWSheet.Range("B2:" + char1 + char2 + (studcount + 1).ToString)

            ' For i As Integer = 0 To studcount - 1 Step 1

            'headarray(i, 0) = Studobjects(i).rollno

            'headarray(i, 1) = Studobjects(i).sname

            'Dim k As Integer = 2

            'For j As Integer = 0 To Studobjects(0).scount - 1 Step 1

            'headarray(i, k) = (Integer.Parse(Studobjects(i).SubjectC(j)) + Integer.Parse(Studobjects(i).SubjectA(j)))

            'k = k + 1

            'headarray(i, k) = Studobjects(i).SubjectE(j)

            'k = k + 1

            'headarray(i, k) = Studobjects(i).SubjectT(j)

            'k = k + 1

            'Next

            'Next

            Dim testarray(10, 10) As String

            For i = 0 To 10 Step 1

                For j = 0 To 10 Step 1

                    testarray(i, j) = " "

                Next

            Next

            testarray(1, 0) = "No. of students who appeared for the test"

            testarray(1, 1) = Label10.Text.ToString

            testarray(2, 0) = "No. of students who passed"

            testarray(2, 1) = Label11.Text.ToString

            testarray(3, 0) = "No. of students who failed"

            testarray(3, 1) = Label12.Text.ToString

            testarray(4, 0) = "Pass Percentage"

            testarray(4, 1) = Label13.Text.ToString

            testarray(5, 0) = "Fail Percentage"

            testarray(5, 1) = Label14.Text.ToString

            testarray(6, 0) = "Maximum Marks"

            testarray(7, 0) = "Minimum Marks"

            testarray(0, 3) = Label40.Text.ToString

            testarray(0, 4) = Label41.Text.ToString

            testarray(0, 5) = Label42.Text.ToString

            testarray(0, 6) = Label43.Text.ToString

            testarray(0, 7) = Label44.Text.ToString

            testarray(0, 8) = Label45.Text.ToString

            testarray(0, 9) = Label46.Text.ToString

            testarray(0, 10) = Label47.Text.ToString

            testarray(6, 3) = Label23.Text.ToString

            testarray(6, 4) = Label24.Text.ToString

            testarray(6, 5) = Label25.Text.ToString

            testarray(6, 6) = Label26.Text.ToString

            testarray(6, 7) = Label27.Text.ToString

            testarray(6, 8) = Label28.Text.ToString

            testarray(6, 9) = Label29.Text.ToString

            testarray(6, 10) = Label30.Text.ToString

            testarray(7, 3) = Label15.Text.ToString

            testarray(7, 4) = Label16.Text.ToString

            testarray(7, 5) = Label17.Text.ToString

            testarray(7, 6) = Label18.Text.ToString

            testarray(7, 7) = Label19.Text.ToString

            testarray(7, 8) = Label20.Text.ToString

            testarray(7, 9) = Label21.Text.ToString

            testarray(7, 10) = Label22.Text.ToString

            Dim piechart1 As Microsoft.Office.Interop.Excel.Chart

            piechart1 = ExWBook.Charts.Add()

            ExRange.Value = testarray

            ExRange.Columns.AutoFit()
            'ExWSheet.Range("B2:B6").
            'ExWBook.ActiveChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie

            'ExWBook.ActiveChart.SetSourceData(ExWSheet.Range("A3:B4"), Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)

            '            ExWBook.ActiveChart.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsNewSheet, )

            'ExRange.Value = headarray

            ExWBook.SaveAs(SaveFileDialog1.FileName)

            ExWBook.Close()

            ExApp.Quit()



        End If

    End Sub

  
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        AboutBox2.Show()

    End Sub
End Class