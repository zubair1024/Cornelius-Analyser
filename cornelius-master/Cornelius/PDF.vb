Option Strict On
Option Explicit On

Imports iTextSharp.text.pdf

Public Class PdfManipulation2

    Private Class MyPdfReader
        Inherits iTextSharp.text.pdf.PdfReader

        Private Sub New(ByVal filename As String)
            MyBase.New(filename)
            Me.encrypted = False
        End Sub

        Private Sub New(ByVal filename As String, ByVal password As String)
            MyBase.New(filename, PdfWriter.GetISOBytes(password))
            Me.encrypted = False
        End Sub

        Friend Overloads Shared Function GetInstance(ByVal filePath As String) As MyPdfReader
            Return New MyPdfReader(filePath)
        End Function

        Friend Overloads Shared Function GetInstance(ByVal filePath As String, ByVal password As String) As MyPdfReader
            Return New MyPdfReader(filePath, password)
        End Function
    End Class

    Private Class DocumentEx
        Inherits iTextSharp.text.Document

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(ByVal pageSize As iTextSharp.text.Rectangle)
            MyBase.New(pageSize)
        End Sub

        Public Sub New(ByVal pageSize As iTextSharp.text.Rectangle, _
                       ByVal leftMargin As Single, ByVal rightMargin As Single, _
                       ByVal topMargin As Single, ByVal bottomMargin As Single)
            MyBase.New(pageSize, leftMargin, rightMargin, topMargin, bottomMargin)
        End Sub

        Public Overloads Function AddProducer(ByVal producerName As String) As Boolean
            Return Me.Add(New iTextSharp.text.Meta(iTextSharp.text.Element.PRODUCER, producerName))
        End Function
    End Class

    ''' <summary>
    ''' Remove all resttrictions from a pdf file
    ''' </summary>
    ''' <param name="restrictedPdf">The full path to the restricted pdf file</param>
    ''' <param name="password">Requires only if the restricted pdf is password protected.</param>
    ''' <param name="saveABackup">If True, the original restricted pdf will be saved as [filename]_BAK.pdf. Else, it will be overwritten.</param>
    ''' <returns>True if the operation succeeded. False otherwise</returns>
    ''' <remarks></remarks>
    Public Shared Function RemoveRestrictions(ByVal restrictedPdf As String, Optional ByVal password As String = Nothing, Optional ByVal saveABackup As Boolean = True) As Boolean
        Dim result As Boolean = True
        Try
            Dim outputPdf As String = String.Format("{0}\{1}.{2}", IO.Path.GetDirectoryName(restrictedPdf), Date.Now.ToString("yyyyMMddHHmmss"), "pdf")
            Dim reader As MyPdfReader = Nothing
            If String.IsNullOrEmpty(password) Then
                reader = MyPdfReader.GetInstance(restrictedPdf)
            Else
                reader = MyPdfReader.GetInstance(restrictedPdf, password)
            End If
            'create a filestream for output
            Dim fs As New System.IO.FileStream(outputPdf, IO.FileMode.Create, IO.FileAccess.Write)
            'use stamper to copy the source pdf to output
            Dim stamper As New iTextSharp.text.pdf.PdfStamper(reader, fs)
            'remove restrictions
            Dim perms As Integer = iTextSharp.text.pdf.PdfWriter.ALLOW_ASSEMBLY Or _
                                   iTextSharp.text.pdf.PdfWriter.ALLOW_COPY Or _
                                   iTextSharp.text.pdf.PdfWriter.ALLOW_DEGRADED_PRINTING Or _
                                   iTextSharp.text.pdf.PdfWriter.ALLOW_FILL_IN Or _
                                   iTextSharp.text.pdf.PdfWriter.ALLOW_MODIFY_ANNOTATIONS Or _
                                   iTextSharp.text.pdf.PdfWriter.ALLOW_MODIFY_CONTENTS Or _
                                   iTextSharp.text.pdf.PdfWriter.ALLOW_PRINTING Or _
                                   iTextSharp.text.pdf.PdfWriter.ALLOW_SCREENREADERS

            stamper.SetEncryption(False, "", "", perms)
            stamper.Close()
            fs.Dispose()
            reader.Close()
            If saveABackup = True Then
                My.Computer.FileSystem.RenameFile(restrictedPdf, String.Format("{0}_{1}.{2}", IO.Path.GetFileNameWithoutExtension(restrictedPdf), "BAK", "pdf"))
            Else
                IO.File.Delete(restrictedPdf)
            End If
            My.Computer.FileSystem.RenameFile(outputPdf, IO.Path.GetFileName(restrictedPdf))
        Catch ex As Exception
            Debug.Write(ex.Message)
            result = False
        End Try
        Return result
    End Function

    ''' <summary>
    ''' Extract the text from pdf pages and return it as a string
    ''' </summary>
    ''' <param name="sourcePDF">Full path to the source pdf file</param>
    ''' <param name="fromPageNum">[Optional] the page number (inclusive) to start text extraction </param>
    ''' <param name="toPageNum">[Optional] the page number (inclusive) to stop text extraction</param>
    ''' <returns>A string containing the text extracted from the specified pages</returns>
    ''' <remarks>If fromPageNum is not specified, text extraction will start from page 1. If
    ''' toPageNum is not specified, text extraction will end at the last page of the source pdf file.</remarks>
    Public Shared Function ParsePdfText(ByVal sourcePDF As String, _
                                  Optional ByVal fromPageNum As Integer = 0, _
                                  Optional ByVal toPageNum As Integer = 0) As String

        Dim sb As New System.Text.StringBuilder()
        Try
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePDF)
            Dim pageBytes() As Byte = Nothing
            Dim token As iTextSharp.text.pdf.PRTokeniser = Nothing
            Dim tknType As Integer = -1
            Dim tknValue As String = String.Empty

            If fromPageNum = 0 Then
                fromPageNum = 1
            End If
            If toPageNum = 0 Then
                toPageNum = reader.NumberOfPages
            End If

            If fromPageNum > toPageNum Then
                Throw New ApplicationException("Parameter error: The value of fromPageNum can " & _
                                           "not be larger than the value of toPageNum")
            End If

            For i As Integer = fromPageNum To toPageNum Step 1
                pageBytes = reader.GetPageContent(i)
                If Not IsNothing(pageBytes) Then
                    token = New iTextSharp.text.pdf.PRTokeniser(pageBytes)
                    While token.NextToken()
                        tknType = token.TokenType()
                        tknValue = token.StringValue
                        Select Case tknType
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.NUMBER      '1
                                Dim dValue As Double
                                If Double.TryParse(tknValue, dValue) Then
                                    If dValue < -8000 Then
                                        sb.Append(ControlChars.Tab)
                                    End If
                                End If
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.STRING      '2
                                sb.Append(token.StringValue)
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.NAME        '3
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.COMMENT     '4
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.START_ARRAY '5
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.END_ARRAY   '6
                                sb.Append(" ")
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.START_DIC   '7
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.END_DIC     '8
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.REF         '9
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.OTHER       '10
                                Select Case tknValue
                                    Case "TJ"
                                        sb.Append(" ")
                                    Case "ET", "TD", "Td", "Tm", "T*"
                                        sb.Append(Environment.NewLine)
                                End Select
                        End Select
                    End While
                End If
            Next i
            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Exception occured. " & ex.Message)
            Return String.Empty
        End Try
        Return sb.ToString()
    End Function

    Public Shared Function ParseAllPdfText(ByVal sourcePDF As String) As Dictionary(Of Integer, String)
        Dim pdfText As New Dictionary(Of Integer, String)
        Dim sb As New System.Text.StringBuilder()
        Try
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePDF)
            Dim pageBytes() As Byte = Nothing
            Dim token As iTextSharp.text.pdf.PRTokeniser = Nothing
            Dim tknType As Integer = -1
            Dim tknValue As String = String.Empty

            For i As Integer = 1 To reader.NumberOfPages Step 1
                pageBytes = reader.GetPageContent(i)
                If Not IsNothing(pageBytes) Then
                    sb.Length = 0
                    token = New iTextSharp.text.pdf.PRTokeniser(pageBytes)
                    While token.NextToken()
                        tknType = token.TokenType()
                        tknValue = token.StringValue
                        Select Case tknType
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.NUMBER      '1
                                Dim dValue As Double
                                If Double.TryParse(tknValue, dValue) Then
                                    If dValue < -8000 Then
                                        sb.Append(ControlChars.Tab)
                                    End If
                                End If
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.STRING      '2
                                sb.Append(token.StringValue)
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.NAME        '3
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.COMMENT     '4
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.START_ARRAY '5
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.END_ARRAY   '6
                                sb.Append(" ")
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.START_DIC   '7
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.END_DIC     '8
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.REF         '9
                                'Ignore
                            Case iTextSharp.text.pdf.PRTokeniser.TokType.OTHER       '10
                                Select Case tknValue
                                    Case "TJ"
                                        sb.Append(" ")
                                    Case "ET", "TD", "Td", "Tm", "T*"
                                        sb.Append(Environment.NewLine)
                                End Select
                        End Select
                    End While
                    pdfText.Add(i, sb.ToString)
                End If
            Next i
            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Exception occured. " & ex.Message)
        End Try
        Return pdfText
    End Function

    ''' <summary>
    ''' Textually compare 2 pdf files page by page and write the difference to a text file.
    ''' </summary>
    ''' <param name="pdf1">the full path to 1st pdf file</param>
    ''' <param name="pdf2">the full path to 2nd pdf file</param>
    ''' <param name="resultFile">the full path to the result file</param>
    ''' <param name="fromPageNum">page number to start comparing</param>
    ''' <param name="toPageNum">page number to stop comparing</param>
    ''' <remarks>If no values are specified for fromPageNum and toPageNum, the sub will
    ''' compare every page in the input pdfs.</remarks>
    Public Shared Sub ComparePdfs(ByVal pdf1 As String, ByVal pdf2 As String, _
                                  ByVal resultFile As String, _
                                  Optional ByVal fromPageNum As Integer = 0, _
                                  Optional ByVal toPageNum As Integer = 0)
        Try
            'For pdf1
            Dim reader1 As New iTextSharp.text.pdf.PdfReader(pdf1)
            Dim pageCount1 As Integer = reader1.NumberOfPages
            Dim pageBytes1() As Byte = Nothing
            Dim token1 As iTextSharp.text.pdf.PRTokeniser = Nothing
            Dim tknType1 As Integer = -1
            Dim tknValue1 As String = String.Empty

            'For pdf2
            Dim reader2 As New iTextSharp.text.pdf.PdfReader(pdf2)
            Dim pageCount2 As Integer = reader2.NumberOfPages
            Dim pageBytes2() As Byte = Nothing
            Dim token2 As iTextSharp.text.pdf.PRTokeniser = Nothing
            Dim tknType2 As Integer = -1
            Dim tknValue2 As String = String.Empty

            If fromPageNum = 0 Then
                fromPageNum = 1
            End If

            If toPageNum = 0 Then
                toPageNum = Math.Min(pageCount1, pageCount2)
            Else
                If toPageNum > pageCount1 OrElse toPageNum > pageCount2 Then
                    toPageNum = Math.Min(pageCount1, pageCount2)
                End If
            End If

            If fromPageNum > toPageNum Then
                Throw New ApplicationException("Parameter error: The value of fromPageNum can " & _
                                           "not be larger than the value of toPageNum")
            End If

            Dim writer As New System.IO.StreamWriter(resultFile)
            For i As Integer = fromPageNum To toPageNum Step 1
                writer.WriteLine("Differences found in page " & i)
                pageBytes1 = reader1.GetPageContent(i)
                pageBytes2 = reader2.GetPageContent(i)
                If Not IsNothing(pageBytes1) AndAlso Not IsNothing(pageBytes2) Then
                    token1 = New iTextSharp.text.pdf.PRTokeniser(pageBytes1)
                    token2 = New iTextSharp.text.pdf.PRTokeniser(pageBytes2)
                    While token1.NextToken() AndAlso token2.NextToken()

                        tknType1 = token1.TokenType()
                        tknValue1 = token1.StringValue

                        tknType2 = token2.TokenType()
                        tknValue2 = token2.StringValue

                        If tknType1 = iTextSharp.text.pdf.PRTokeniser.TokType.STRING AndAlso _
                           tknType2 = iTextSharp.text.pdf.PRTokeniser.TokType.STRING Then
                            If String.Compare(tknValue1, tknValue2) <> 0 Then
                                writer.WriteLine("Pdf1: " & tknValue1 & " <> Pdf2: " & tknValue2)
                            End If
                        End If
                    End While
                End If
            Next i
            writer.Close()
            reader1.Close()
            reader2.Close()
        Catch ex As Exception
            MessageBox.Show("Exception occured. " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Extract a single page from source pdf to a new pdf
    ''' </summary>
    ''' <param name="sourcePdf">the full path to source pdf file</param>
    ''' <param name="pageNumberToExtract">the page number to extract</param>
    ''' <param name="outPdf">the full path for the output pdf</param>
    ''' <remarks></remarks>
    Public Overloads Shared Sub ExtractPdfPage(ByVal sourcePdf As String, ByVal pageNumberToExtract As Integer, ByVal outPdf As String)
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        'Dim doc As PdfManipulation.DocumentEx = Nothing
        Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
            'doc = New PdfManipulation.DocumentEx(reader.GetPageSizeWithRotation(pageNumberToExtract))
            'Debug.WriteLine("Add producer: " & doc.AddProducer().ToString)
            pdfCpy = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outPdf, IO.FileMode.Create))
            doc.Open()
            page = pdfCpy.GetImportedPage(reader, pageNumberToExtract)
            pdfCpy.AddPage(page)
            doc.Close()
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Extract selected pages from a source pdf to a new pdf
    ''' </summary>
    ''' <param name="sourcePdf">the full path to source pdf to a new pdf</param>
    ''' <param name="pageNumbersToExtract">the page numbers to extract (i.e {1, 3, 5, 6})</param>
    ''' <param name="outPdf">The full path for the output pdf</param>
    ''' <remarks>The output pdf will contains the extracted pages in the order of the page numbers listed
    ''' in pageNumbersToExtract parameter.</remarks>
    Public Overloads Shared Sub ExtractPdfPage(ByVal sourcePdf As String, ByVal pageNumbersToExtract As Integer(), ByVal outPdf As String)
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
            pdfCpy = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outPdf, IO.FileMode.Create))
            doc.Open()
            For Each pageNum As Integer In pageNumbersToExtract
                page = pdfCpy.GetImportedPage(reader, pageNum)
                pdfCpy.AddPage(page)
            Next
            doc.Close()
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    ''' <summary>
    ''' Extract pages from an existing pdf file to create a new pdf with bookmarks preserved
    ''' </summary>
    ''' <param name="sourcePdf">full path to sthe source pdf</param>
    ''' <param name="pageNumbersToExtract">an integer array containing the page number of the pages to be extracted</param>
    ''' <param name="outPdf">the full path to the output pdf</param>
    ''' <remarks></remarks>
    Public Shared Sub ExtractPdfPages(ByVal sourcePdf As String, ByVal pageNumbersToExtract As List(Of Integer), ByVal outPdf As String)

        Dim raf As iTextSharp.text.pdf.RandomAccessFileOrArray = Nothing
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim outlines As IList(Of Dictionary(Of String, Object)) = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Dim stamper As iTextSharp.text.pdf.PdfStamper = Nothing
        Dim hshTable As System.Collections.Hashtable = Nothing
        Try
            raf = New iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf)
            reader = New iTextSharp.text.pdf.PdfReader(raf, Nothing)
            outlines = iTextSharp.text.pdf.SimpleBookmark.GetBookmark(reader)
            reader.SelectPages(pageNumbersToExtract)
            stamper = New iTextSharp.text.pdf.PdfStamper(reader, New IO.FileStream(outPdf, IO.FileMode.Create))
            RemoveUnusedBookmarks(outlines, pageNumbersToExtract)
            stamper.Outlines = outlines
            stamper.Close()
            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Shared Sub RemoveUnusedBookmarks(ByRef bookmarks As IList(Of Dictionary(Of String, Object)), ByVal pagesToKeep As List(Of Integer))
        Dim bookmark As Dictionary(Of String, Object) = Nothing
        Dim obj As Object = Nothing
        For i As Integer = bookmarks.Count - 1 To 0 Step -1
            obj = bookmarks(i)
            If TypeOf obj Is IList(Of Dictionary(Of String, Object)) Then
                RemoveUnusedBookmarks(DirectCast(obj, IList(Of Dictionary(Of String, Object))), pagesToKeep)
            ElseIf TypeOf obj Is Dictionary(Of String, Object) Then
                bookmark = DirectCast(obj, Dictionary(Of String, Object))
                If bookmark.ContainsKey("Page") Then
                    Dim value As String = DirectCast(bookmark.Item("Page"), String)
                    If Not String.IsNullOrEmpty(value) Then
                        Dim parts() As String = value.Split(" "c)
                        If parts.Length > 0 Then
                            Dim pageNum As Integer = -1
                            If Integer.TryParse(parts(0), pageNum) Then
                                Dim idx As Integer = pagesToKeep.IndexOf(pageNum)
                                If idx < 0 Then
                                    bookmarks.Remove(bookmark)
                                Else
                                    parts(0) = (idx + 1).ToString
                                    value = String.Join(" ", parts)
                                    bookmark.Item("Page") = value
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub


    ''' <summary>
    ''' Split a single pdf file into multiple pdfs with equal number of pages.
    ''' </summary>
    ''' <param name="sourcePdf">the full path to the source pdf</param>
    ''' <param name="parts">the number of splitted pdfs to split to</param>
    ''' <param name="baseNameOutPdf">the base file name (full path) for splitted pdfs.
    ''' The actual output pdf file names will be serialized. </param>
    ''' <remarks>The last splitted pdf may not have
    ''' the same number of pages as the rest, depending on the combination of number of pages in the source pdf 
    ''' and the number of parts to be splitted. For example, if the original pdf has 9 pages and it is to be 
    ''' splitted into 5 parts, the last splitted pdf will have only 1 page while all others have 2 pages.</remarks>
    Public Shared Sub SplitPdfByParts(ByVal sourcePdf As String, ByVal parts As Integer, ByVal baseNameOutPdf As String)
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Dim pageCount As Integer = 0
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            pageCount = reader.NumberOfPages
            If pageCount < parts Then
                Throw New ArgumentException("Not enough pages in source pdf to split")
            Else
                Dim n As Integer = pageCount \ parts
                Dim currentPage As Integer = 1
                Dim ext As String = IO.Path.GetExtension(baseNameOutPdf)
                Dim outfile As String = String.Empty
                For i As Integer = 1 To parts
                    outfile = baseNameOutPdf.Replace(ext, "_" & i & ext)
                    doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(currentPage))
                    pdfCpy = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outfile, IO.FileMode.Create))
                    doc.Open()
                    If i < parts Then
                        For j As Integer = 1 To n
                            page = pdfCpy.GetImportedPage(reader, currentPage)
                            pdfCpy.AddPage(page)
                            currentPage += 1
                        Next j
                    Else
                        For j As Integer = currentPage To pageCount
                            page = pdfCpy.GetImportedPage(reader, j)
                            pdfCpy.AddPage(page)
                        Next j
                    End If
                    doc.Close()
                Next
            End If
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub



    ''' <summary>
    ''' Split source pdf into multiple pdfs with specifc number of pages
    ''' </summary>
    ''' <param name="sourcePdf">the full path to source pdf</param>
    ''' <param name="numOfPages">the number of pages each splitted pdf should contain</param>
    ''' <param name="baseNameOutPdf">the base file name (full path) for splitted pdfs.
    ''' The actual output pdf file names will be serialized. </param>
    ''' <remarks>The last splitted pdf may not have
    ''' the same number of pages as the rest, depending on the combination of number of pages in the source pdf 
    ''' and the number of target pages in each splitted pdf. For example, if the original pdf has 9 pages and it is to be 
    ''' splitted with 2 pages for each pdf, the last splitted pdf will have only 1 page while all others have 2 pages.</remarks>
    Public Shared Sub SplitPdfByPages(ByVal sourcePdf As String, ByVal numOfPages As Integer, ByVal baseNameOutPdf As String)
        Dim raf As iTextSharp.text.pdf.RandomAccessFileOrArray = Nothing
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Dim pageCount As Integer = 0

        Try
            raf = New iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf)
            reader = New iTextSharp.text.pdf.PdfReader(raf, Nothing)
            pageCount = reader.NumberOfPages
            If pageCount < numOfPages Then
                Throw New ArgumentException("Not enough pages in source pdf to split")
            Else
                Dim ext As String = IO.Path.GetExtension(baseNameOutPdf)
                Dim outfile As String = String.Empty
                Dim n As Integer = CInt(Math.Ceiling(pageCount / numOfPages))
                Dim currentPage As Integer = 1
                For i As Integer = 1 To n
                    outfile = baseNameOutPdf.Replace(ext, "_" & i & ext)
                    doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(currentPage))
                    pdfCpy = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outfile, IO.FileMode.Create))
                    doc.Open()
                    If i < n Then
                        For j As Integer = 1 To numOfPages
                            page = pdfCpy.GetImportedPage(reader, currentPage)
                            pdfCpy.AddPage(page)
                            currentPage += 1
                        Next j
                    Else
                        For j As Integer = currentPage To pageCount
                            page = pdfCpy.GetImportedPage(reader, j)
                            pdfCpy.AddPage(page)
                        Next j
                    End If
                    doc.Close()
                Next
            End If
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Extract pages from multiple pdf's file and merge them into 
    ''' a single pdf
    ''' </summary>
    ''' <param name="sourceTable">the datatable containing source pfd paths and the pages to extract
    ''' from each of them. This datatable should have 2 datacolumns of type String. The 1st column (column 0)
    ''' is for the file (full) path while the 2nd column (column 1) is for the list of pages to extract from
    ''' the source pdf in column 1. This list is a string of integer values separated by commas 
    ''' (ex: "1, 3, 2, 5 , 8, 7, 9") </param>
    ''' <param name="outPdf">the path to save the output pdf</param>
    ''' <remarks>the pdf pages are extracted and merged in the order listed in the source datatable.
    ''' That is, for source pdf files, they will be merged from top row down, and for pages, they will be merged
    ''' by the order listed in the csv string</remarks>
    Public Shared Sub ExtractAndMergePdfPages(ByVal sourceTable As DataTable, ByVal outPdf As String)
        Dim rowCount As Integer = sourceTable.Rows.Count
        Dim sourcePdf As String = String.Empty
        Dim pageNumbersToExtract() As Integer = Nothing
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Select Case rowCount
            Case 0  'Nothing to extract and merge
                Exit Sub
            Case 1  'only 1 source pdf
                sourcePdf = CStr(sourceTable.Rows(0).Item(0))
                pageNumbersToExtract = ConvertToIntegerArray(CStr(sourceTable.Rows(0).Item(1)))
                ExtractPdfPage(sourcePdf, pageNumbersToExtract, outPdf)
            Case Else   'multiple source pdf's
                Try
                    sourcePdf = CStr(sourceTable.Rows(0).Item(0))
                    pageNumbersToExtract = ConvertToIntegerArray(CStr(sourceTable.Rows(0).Item(1)))
                    reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
                    doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
                    pdfCpy = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outPdf, IO.FileMode.Create))
                    doc.Open()
                    For Each pageNum As Integer In pageNumbersToExtract
                        page = pdfCpy.GetImportedPage(reader, pageNum)
                        pdfCpy.AddPage(page)
                    Next
                    reader.Close()
                    For i As Integer = 1 To rowCount - 1
                        sourcePdf = CStr(sourceTable.Rows(i).Item(0))
                        pageNumbersToExtract = ConvertToIntegerArray(CStr(sourceTable.Rows(i).Item(1)))
                        reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
                        doc.SetPageSize(reader.GetPageSizeWithRotation(1))
                        For Each pageNum As Integer In pageNumbersToExtract
                            page = pdfCpy.GetImportedPage(reader, pageNum)
                            pdfCpy.AddPage(page)
                        Next
                        reader.Close()
                    Next
                    doc.Close()
                Catch ex As Exception
                    Throw ex
                End Try
        End Select
    End Sub

    ''' <summary>
    ''' Helper function to convert a csv integer string to an integer array
    ''' </summary>
    ''' <param name="csvNumbers">the integer string in csv format (ex: "1, 5, 7, 4")</param>
    ''' <returns>Integer array converted from the csv string (ex: {1, 5, 7, 4}</returns>
    ''' <remarks>No error checking/handling. If the input string contains non-numeric values
    ''' the function will crash. It's up to you to handle this error.</remarks>
    Private Shared Function ConvertToIntegerArray(ByVal csvNumbers As String) As Integer()
        Dim numbers() As String = csvNumbers.Split(",".ToCharArray, System.StringSplitOptions.RemoveEmptyEntries)
        Dim upperBound As Integer = numbers.Length - 1
        Dim output(upperBound) As Integer
        For i As Integer = 0 To upperBound
            output(i) = Integer.Parse(numbers(i))
        Next
        Return output
    End Function

    Private Shared Sub SetSecurityPasswords(ByVal sourcePdf As String, ByVal outputPdf As String, ByVal userPassword As String, ByVal ownerPassword As String)
        Try
            Dim outStream As New IO.FileStream(outputPdf, IO.FileMode.Create)
            Dim reader As New PdfReader(sourcePdf)
            Dim userPwdBytes() As Byte = System.Text.Encoding.ASCII.GetBytes(userPassword)
            Dim ownerPwdBytes() As Byte = System.Text.Encoding.ASCII.GetBytes(ownerPassword)
            Dim permissions As Integer = PdfWriter.ALLOW_DEGRADED_PRINTING Or PdfWriter.ALLOW_COPY
            PdfEncryptor.Encrypt(reader, outStream, userPwdBytes, ownerPwdBytes, permissions, PdfWriter.STRENGTH128BITS)
            reader.Close()
        Catch ex As Exception
            'Put your own code for exception handling here

        End Try
    End Sub

    ''' <summary>
    ''' Add and image as the watermark on each page of the source pdf to create a new pdf with watermark
    ''' </summary>
    ''' <param name="sourceFile">the full path to the source pdf</param>
    ''' <param name="outputFile">the full path where the watermarked pdf will be saved to</param>
    ''' <param name="watermarkImage">the full path to the image file to use as the watermark</param>
    ''' <remarks>The watermark image will be align in the center of each page</remarks>
    Public Shared Sub AddWatermarkImage(ByVal sourceFile As String, ByVal outputFile As String, ByVal watermarkImage As String)
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim stamper As iTextSharp.text.pdf.PdfStamper = Nothing
        Dim img As iTextSharp.text.Image = Nothing
        Dim cb As iTextSharp.text.pdf.PdfContentByte = Nothing
        Dim rect As iTextSharp.text.Rectangle = Nothing
        Dim X, Y As Single
        Dim pageCount As Integer = 0
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourceFile)
            rect = reader.GetPageSizeWithRotation(1)
            stamper = New iTextSharp.text.pdf.PdfStamper(reader, New System.IO.FileStream(outputFile, IO.FileMode.Create))
            img = iTextSharp.text.Image.GetInstance(watermarkImage)
            If img.Width > rect.Width OrElse img.Height > rect.Height Then
                img.ScaleToFit(rect.Width, rect.Height)
                X = (rect.Width - img.ScaledWidth) / 2
                Y = (rect.Height - img.ScaledHeight) / 2
            Else
                X = (rect.Width - img.Width) / 2
                Y = (rect.Height - img.Height) / 2
            End If
            img.SetAbsolutePosition(X, Y)
            pageCount = reader.NumberOfPages()
            For i As Integer = 1 To pageCount
                cb = stamper.GetUnderContent(i)
                cb.AddImage(img)
            Next
            stamper.Close()
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Add text as the watermark to each page of the source pdf to create a new pdf with text watermark
    ''' </summary>
    ''' <param name="sourceFile">the full path to the source pdf file</param>
    ''' <param name="outputFile">the full path where the watermarked pdf file will be saved to</param>
    ''' <param name="watermarkText">the string array conntaining the text to use as the watermark. Each element is treated as a line in the watermark</param>
    ''' <param name="watermarkFont">the font to use for the watermark. The default font is HELVETICA</param>
    ''' <param name="watermarkFontSize">the size of the font. The default size is 48</param>
    ''' <param name="watermarkFontColor">the color of the watermark. The default color is blue</param>
    ''' <param name="watermarkFontOpacity">the opacity of the watermark. The default opacity is 0.3</param>
    ''' <param name="watermarkRotation">the rotation in degree of the watermark. The default rotation is 45 degree</param>
    ''' <remarks></remarks>
    Public Shared Sub AddWatermarkText(ByVal sourceFile As String, ByVal outputFile As String, ByVal watermarkText() As String, _
                                       Optional ByVal watermarkFont As iTextSharp.text.pdf.BaseFont = Nothing, _
                                       Optional ByVal watermarkFontSize As Single = 48, _
                                       Optional ByVal watermarkFontColor As iTextSharp.text.BaseColor = Nothing, _
                                       Optional ByVal watermarkFontOpacity As Single = 0.3F, _
                                       Optional ByVal watermarkRotation As Single = 45.0F)

        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim stamper As iTextSharp.text.pdf.PdfStamper = Nothing
        Dim gstate As iTextSharp.text.pdf.PdfGState = Nothing
        Dim underContent As iTextSharp.text.pdf.PdfContentByte = Nothing
        Dim rect As iTextSharp.text.Rectangle = Nothing
        Dim currentY As Single = 0.0F
        Dim offset As Single = 0.0F
        Dim pageCount As Integer = 0
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourceFile)
            rect = reader.GetPageSizeWithRotation(1)
            stamper = New iTextSharp.text.pdf.PdfStamper(reader, New System.IO.FileStream(outputFile, IO.FileMode.Create))
            If watermarkFont Is Nothing Then
                watermarkFont = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.HELVETICA, _
                                                              iTextSharp.text.pdf.BaseFont.CP1252, _
                                                              iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED)
            End If
            If watermarkFontColor Is Nothing Then
                watermarkFontColor = iTextSharp.text.BaseColor.BLUE
            End If
            gstate = New iTextSharp.text.pdf.PdfGState()
            gstate.FillOpacity = watermarkFontOpacity
            gstate.StrokeOpacity = watermarkFontOpacity
            pageCount = reader.NumberOfPages()
            For i As Integer = 1 To pageCount
                underContent = stamper.GetUnderContent(i)
                With underContent
                    .SaveState()
                    .SetGState(gstate)
                    .SetColorFill(watermarkFontColor)
                    .BeginText()
                    .SetFontAndSize(watermarkFont, watermarkFontSize)
                    .SetTextMatrix(30, 30)
                    If watermarkText.Length > 1 Then
                        currentY = (rect.Height / 2) + ((watermarkFontSize * watermarkText.Length) / 2)
                    Else
                        currentY = (rect.Height / 2)
                    End If
                    For j As Integer = 0 To watermarkText.Length - 1
                        If j > 0 Then
                            offset = (j * watermarkFontSize) + (watermarkFontSize / 4) * j
                        Else
                            offset = 0.0F
                        End If
                        .ShowTextAligned(iTextSharp.text.Element.ALIGN_CENTER, watermarkText(j), rect.Width / 2, currentY - offset, watermarkRotation)
                    Next
                    .EndText()
                    .RestoreState()
                End With
            Next
            stamper.Close()
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Merge multiple pdf files into 1 preserving all bookmarks
    ''' </summary>
    ''' <param name="sourcePdfs">string array containing full path to the source pdf's</param>
    ''' <param name="outputPdf">full path to the output (merged) pdf</param>
    ''' <returns>True if successful. False otherwise.</returns>
    ''' <remarks></remarks>
    Public Shared Function MergePdfFilesWithBookmarks(ByVal sourcePdfs() As String, ByVal outputPdf As String) As Boolean
        Dim result As Boolean = False
        Dim pdfCount As Integer = 0     'total input pdf file count
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pdfDoc As iTextSharp.text.Document = Nothing    'the output pdf document
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim pageCount As Integer = 0    'number of pages in the current pdf
        Dim totalPages As Integer = 0   'number of pages so far in the merged pdf
        Dim bookmarks As New System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of String, Object))
        Dim tempBookmarks As System.Collections.Generic.IList(Of System.Collections.Generic.Dictionary(Of String, Object)) = Nothing
        ' Must have more than 1 source pdf's to merge
        If sourcePdfs.Length > 1 Then
            Try
                For i As Integer = 0 To sourcePdfs.GetUpperBound(0)
                    reader = New iTextSharp.text.pdf.PdfReader(sourcePdfs(i))
                    reader.ConsolidateNamedDestinations()
                    pageCount = reader.NumberOfPages
                    tempBookmarks = SimpleBookmark.GetBookmark(reader)
                    If i = 0 Then
                        pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
                        pdfCpy = New iTextSharp.text.pdf.PdfCopy(pdfDoc, New System.IO.FileStream(outputPdf, IO.FileMode.Create))
                        pdfDoc.Open()
                        totalPages = pageCount
                    Else
                        If tempBookmarks IsNot Nothing Then
                            SimpleBookmark.ShiftPageNumbers(tempBookmarks, totalPages, Nothing)
                        End If
                        totalPages += pageCount
                    End If
                    If tempBookmarks IsNot Nothing Then
                        bookmarks.AddRange(tempBookmarks)
                    End If
                    For n As Integer = 1 To pageCount
                        page = pdfCpy.GetImportedPage(reader, n)
                        pdfCpy.AddPage(page)
                    Next
                    reader.Close()
                Next
                pdfCpy.Outlines = bookmarks
                pdfDoc.Close()
                result = True
            Catch ex As Exception
                Throw New ApplicationException(ex.Message, ex)
            End Try
        End If
        Return result
    End Function

    Public Shared Function ExportBookmarksToXML(ByVal sourcePdf As String, ByVal outputXML As String) As Boolean
        Dim result As Boolean = False
        Try
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePdf)
            Dim bookmarks As System.Collections.Generic.IList(Of System.Collections.Generic.Dictionary(Of String, Object)) = SimpleBookmark.GetBookmark(reader)
            Using outFile As New IO.StreamWriter(outputXML)
                SimpleBookmark.ExportToXML(bookmarks, outFile, "ISO8859-1", True)
            End Using
            reader.Close()
            result = True
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex)
        End Try
        Return result
    End Function

    ''' <summary>
    ''' Merge multiple pdf files into a single pdf
    ''' </summary>
    ''' <param name="pdfFiles">string array containing full paths to the pdf files to be merged</param>
    ''' <param name="outputPath">full path to the merged output pdf</param>
    ''' <param name="authorName">Author's name.</param>
    ''' <param name="creatorName">Creator's name</param>
    ''' <param name="subject">Subject field</param>
    ''' <param name="title">Title field</param>
    ''' <param name="keywords">keywords field</param>
    ''' <returns>True if the merging is successful, False otherwise.</returns>
    ''' <remarks>All optional paramters are used for the output pdf metadata.
    ''' You can see a pdf metada by going to the PDF tab of the file's Property window.</remarks>
    Public Shared Function MergePdfFiles(ByVal pdfFiles() As String, ByVal outputPath As String, _
                                         Optional ByVal authorName As String = "", _
                                         Optional ByVal creatorName As String = "", _
                                         Optional ByVal subject As String = "", _
                                         Optional ByVal title As String = "", _
                                         Optional ByVal keywords As String = "") As Boolean
        Dim result As Boolean = False
        Dim pdfCount As Integer = 0     'total input pdf file count
        Dim f As Integer = 0            'pointer to current input pdf file
        Dim fileName As String = String.Empty   'current input pdf filename
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pageCount As Integer = 0    'cureent input pdf page count
        Dim pdfDoc As iTextSharp.text.Document = Nothing    'the output pdf document
        Dim writer As iTextSharp.text.pdf.PdfWriter = Nothing
        Dim cb As iTextSharp.text.pdf.PdfContentByte = Nothing
        'Declare a variable to hold the imported pages
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Dim rotation As Integer = 0
        'Declare a font to used for the bookmarks
        Dim bookmarkFont As iTextSharp.text.Font = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA, _
                                                                  12, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLUE)
        Try
            pdfCount = pdfFiles.Length
            If pdfCount > 1 Then
                'Open the 1st pad using PdfReader object
                fileName = pdfFiles(f)
                reader = New iTextSharp.text.pdf.PdfReader(fileName)
                'Get page count
                pageCount = reader.NumberOfPages
                'Instantiate an new instance of pdf document and set its margins. This will be the output pdf.
                'NOTE: bookmarks will be added at the 1st page of very original pdf file using its filename. The location
                'of this bookmark will be placed at the upper left hand corner of the document. So you'll need to adjust
                'the margin left and margin top values such that the bookmark won't overlay on the merged pdf page. The 
                'unit used is "points" (72 points = 1 inch), thus in this example, the bookmarks' location is at 1/4 inch from
                'left and 1/4 inch from top of the page.
                pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1), 18, 18, 18, 18)
                'Instantiate a PdfWriter that listens to the pdf document
                writer = iTextSharp.text.pdf.PdfWriter.GetInstance(pdfDoc, New System.IO.FileStream(outputPath, IO.FileMode.Create))
                'Set metadata and open the document
                With pdfDoc
                    .AddAuthor(authorName)
                    .AddCreationDate()
                    .AddCreator(creatorName)
                    .AddProducer()
                    .AddSubject(subject)
                    .AddTitle(title)
                    .AddKeywords(keywords)
                    .Open()
                End With
                'Instantiate a PdfContentByte object
                cb = writer.DirectContent
                'Now loop thru the input pdfs
                While f < pdfCount
                    'Declare a page counter variable
                    Dim i As Integer = 0
                    'Loop thru the current input pdf's pages starting at page 1
                    While i < pageCount
                        i += 1
                        'Get the input page size
                        pdfDoc.SetPageSize(reader.GetPageSizeWithRotation(i))
                        'Create a new page on the output document
                        pdfDoc.NewPage()

                        'If it is the 1st page, we add bookmarks to the page
                        If i = 1 Then
                            'First create a paragraph using the filename as the heading
                            Dim para As New iTextSharp.text.Paragraph(IO.Path.GetFileName(fileName).ToUpper(), bookmarkFont)
                            'Then create a chapter from the above paragraph
                            Dim chpter As New iTextSharp.text.Chapter(para, f + 1)
                            'Finally add the chapter to the document
                            pdfDoc.Add(chpter)
                        End If
                        'Now we get the imported page
                        page = writer.GetImportedPage(reader, i)
                        'Read the imported page's rotation
                        rotation = reader.GetPageRotation(i)
                        'Then add the imported page to the PdfContentByte object as a template based on the page's rotation
                        If rotation = 90 Then
                            cb.AddTemplate(page, 0, -1.0F, 1.0F, 0, 0, reader.GetPageSizeWithRotation(i).Height)
                        ElseIf rotation = 270 Then
                            cb.AddTemplate(page, 0, 1.0F, -1.0F, 0, reader.GetPageSizeWithRotation(i).Width + 60, -30)
                        Else
                            cb.AddTemplate(page, 1.0F, 0, 0, 1.0F, 0, 0)
                        End If
                    End While
                    'Increment f and read the next input pdf file
                    f += 1
                    If f < pdfCount Then
                        fileName = pdfFiles(f)
                        reader = New iTextSharp.text.pdf.PdfReader(fileName)
                        pageCount = reader.NumberOfPages
                    End If
                End While
                'When all done, we close the document so that the pdfwriter object can write it to the output file
                pdfDoc.Close()
                result = True
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return result
    End Function

    Public Shared Sub AddDocumentOutline(ByVal sourcePdf As String, ByVal outputPdf As String, ByVal outlineTable As System.Data.DataTable)
        Dim raf As iTextSharp.text.pdf.RandomAccessFileOrArray = Nothing
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pdfDoc As iTextSharp.text.Document = Nothing
        Dim pdfCpy As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Dim pageCount As Integer = 0

        Dim dv As System.Data.DataView = Nothing
        Dim row As System.Data.DataRow = Nothing

        Try
            'This to ensure that the page number is sorted in ascending order
            With outlineTable
                .Columns(0).ColumnName = "PageNumber"
                .Columns(1).ColumnName = "Title"
                .Columns(2).ColumnName = "MainItem"
            End With
            dv = New System.Data.DataView(outlineTable)
            dv.Sort = "PageNumber ASC"

            raf = New iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf)
            reader = New iTextSharp.text.pdf.PdfReader(raf, Nothing)
            pageCount = reader.NumberOfPages()
            pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
            pdfCpy = New iTextSharp.text.pdf.PdfCopy(pdfDoc, New IO.FileStream(outputPdf, IO.FileMode.Create))
            pdfDoc.Open()

            Dim rootOutline As iTextSharp.text.pdf.PdfOutline = pdfCpy.DirectContent.RootOutline()
            Dim parentOutline As iTextSharp.text.pdf.PdfOutline = Nothing
            Dim childOutline As iTextSharp.text.pdf.PdfOutline = Nothing
            Dim j As Integer = 0
            For pageNum As Integer = 1 To pageCount
                page = pdfCpy.GetImportedPage(reader, pageNum)
                'Check and add outlines
                For curRow As Integer = j To dv.Count - 1
                    row = dv.Item(curRow).Row
                    Dim pageNumber As Integer = CInt(row.Item("PageNumber"))
                    If pageNum = pageNumber Then
                        Dim bookmarkText As String = CStr(row.Item("Title"))
                        Dim isMainItem As Boolean = CBool(row.Item("MainItem"))
                        Dim destination As New iTextSharp.text.pdf.PdfDestination(iTextSharp.text.pdf.PdfDestination.FITH, page.Height)
                        If isMainItem = True Then
                            parentOutline = New iTextSharp.text.pdf.PdfOutline(rootOutline, destination, bookmarkText, False)
                            page.AddOutline(parentOutline, bookmarkText)
                        Else
                            If parentOutline IsNot Nothing Then
                                childOutline = New iTextSharp.text.pdf.PdfOutline(parentOutline, destination, bookmarkText)
                            Else
                                childOutline = New iTextSharp.text.pdf.PdfOutline(rootOutline, destination, bookmarkText)
                            End If
                            page.AddOutline(childOutline, bookmarkText)
                        End If
                    ElseIf pageNumber > pageNum Then
                        j = curRow
                        Exit For
                    End If
                Next
                pdfCpy.AddPage(page)
            Next
            pdfDoc.Close()
            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' This function extract the hyperlinks found on a pdf files.
    ''' </summary>
    ''' <param name="sourcePdf">the full path to the source pdf file</param>
    ''' <param name="pageNumbers">An Integer array containing the page numbers from which the
    ''' the URLs will be extracted. The default value is Nothing, and it will extract URLs from
    ''' the whole document.</param>
    ''' <returns>A datatable containing the URLs and page numbers where they are found</returns>
    ''' <remarks>This function still need more work to extract URLs from Anchor objects or from PRIndirectReference objects.
    ''' I'll will update the code once I found a way to do so</remarks>
    Public Shared Function ExtractURLs(ByVal sourcePdf As String, Optional ByVal pageNumbers() As Integer = Nothing) As System.Data.DataTable
        'We first build a datatable to return the extracted URLs (if any)
        Dim linkTable As New DataTable("ExtractedHyperlinks")
        With linkTable.Columns
            .Add("FoundOnPage", GetType(Integer))
            .Add("URL", GetType(String))
        End With
        Dim row As System.Data.DataRow = Nothing

        'Declare variables
        Dim raf As iTextSharp.text.pdf.RandomAccessFileOrArray = Nothing
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim linkArray As System.Collections.ArrayList = Nothing
        Dim pageDict As iTextSharp.text.pdf.PdfDictionary = Nothing
        Dim pageCount As Integer = 0

        Try
            'Open the pdf file and get page count
            raf = New iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf)
            reader = New iTextSharp.text.pdf.PdfReader(raf, Nothing)
            pageCount = reader.NumberOfPages()

            'Create pageNumbers array if the user did not pass in one
            If pageNumbers Is Nothing Then
                pageNumbers = New Integer(pageCount - 1) {}
                For i As Integer = 0 To pageNumbers.GetUpperBound(0)
                    pageNumbers(i) = i + 1
                Next
            End If

            'We now loop thru the pageNUmbers array to get the urls on each page
            For k As Integer = 0 To pageNumbers.GetUpperBound(0)
                'Get the page dictionary
                Dim page As PdfDictionary = reader.GetPageNRelease(pageNumbers(k))
                'Get the annotation array
                Dim annots As PdfArray = DirectCast(PdfReader.GetPdfObject(page.[Get](PdfName.ANNOTS), page), PdfArray)
                If Not annots Is Nothing Then
                    Dim arr As List(Of iTextSharp.text.pdf.PdfObject) = annots.ArrayList
                    'Now loop thru the annotation arraylist
                    For j As Integer = 0 To arr.Count - 1
                        Dim annoto As PdfObject = PdfReader.GetPdfObject(CType(arr(j), PdfObject))
                        'First we check this PdfObject to make sure that it is a dictionary
                        If TypeOf annoto Is PdfDictionary Then
                            Dim annot As PdfDictionary = DirectCast(annoto, PdfDictionary)
                            'We then get the subtype name and check to see if it's a link
                            If (PdfName.LINK).Equals(annot.Get(PdfName.SUBTYPE)) Then
                                'We now try to get the A name
                                Dim A As PdfObject = annot.Get(PdfName.A)
                                If Not A Is Nothing Then
                                    'We then test to see what type this A name is
                                    If TypeOf A Is PRIndirectReference Then
                                        Dim prIndRef As PRIndirectReference = DirectCast(A, PRIndirectReference)
                                        'Still need work to pull the url from PRIndirectReference object
                                        MsgBox(prIndRef.ToString)
                                    Else
                                        'We again has to make sure the A name is a dictionary
                                        If A.IsDictionary Then
                                            Try
                                                'And finally we try to read the URL from this A name
                                                Dim linkDict As PdfDictionary = CType(A, PdfDictionary)
                                                If linkDict.Contains(PdfName.URI) Then
                                                    'And add the URL to our datatable
                                                    row = linkTable.NewRow()
                                                    row("FoundOnPage") = pageNumbers(k)
                                                    row("URL") = linkDict.Get(PdfName.URI).ToString
                                                    linkTable.Rows.Add(row)
                                                End If
                                            Catch ex As Exception
                                                'Put your code to handle exception here
                                                '
                                            End Try
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Next
            'Close the reader when done to realease resources.
            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Return linkTable
    End Function

    Public Shared Function ExtractImages(ByVal sourcePdf As String) As List(Of Image)
        Dim imgList As New List(Of Image)

        Dim raf As iTextSharp.text.pdf.RandomAccessFileOrArray = Nothing
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pdfObj As iTextSharp.text.pdf.PdfObject = Nothing
        Dim pdfStrem As iTextSharp.text.pdf.PdfStream = Nothing

        Try
            raf = New iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf)
            reader = New iTextSharp.text.pdf.PdfReader(raf, Nothing)
            For i As Integer = 0 To reader.XrefSize - 1
                pdfObj = reader.GetPdfObject(i)
                If Not IsNothing(pdfObj) AndAlso pdfObj.IsStream() Then
                    pdfStrem = DirectCast(pdfObj, iTextSharp.text.pdf.PdfStream)
                    Dim subtype As iTextSharp.text.pdf.PdfObject = pdfStrem.Get(iTextSharp.text.pdf.PdfName.SUBTYPE)
                    If Not IsNothing(subtype) AndAlso subtype.ToString = iTextSharp.text.pdf.PdfName.IMAGE.ToString Then
                        Dim bytes() As Byte = iTextSharp.text.pdf.PdfReader.GetStreamBytesRaw(CType(pdfStrem, iTextSharp.text.pdf.PRStream))
                        If Not IsNothing(bytes) Then
                            Try
                                Using memStream As New System.IO.MemoryStream(bytes)
                                    memStream.Position = 0
                                    Dim img As Image = Image.FromStream(memStream)
                                    imgList.Add(img)
                                End Using
                            Catch ex As Exception
                                'Most likely the image is in an unsupported format
                                'Do nothing
                                'You can add your own code to handle this exception if you want to
                            End Try
                        End If
                    End If
                End If
            Next
            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return imgList
    End Function

    ''' <summary>
    ''' Fill in an AcroForm
    ''' </summary>
    ''' <param name="sourcePdf">the full path to the pdf form file</param>
    ''' <param name="fieldData">a datarow where the column names are the field names in the pdf,
    ''' and the value for that field is the value of the cell in that column</param>
    ''' <param name="outputPdf">the full path of the output file</param>
    ''' <remarks></remarks>
    Public Shared Sub FillAcroForm(ByVal sourcePdf As String, ByVal fieldData As DataRow, ByVal outputPdf As String)
        Try
            'Open the pdf using pdfreader
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePdf)
            'create a filestream for output
            Dim fs As New System.IO.FileStream(outputPdf, IO.FileMode.Create, IO.FileAccess.Write)
            'use stamper to copy the source pdf to output
            Dim stamper As New iTextSharp.text.pdf.PdfStamper(reader, fs)
            'Get the form from the pdf
            Dim frm As iTextSharp.text.pdf.AcroFields = stamper.AcroFields
            'get the fields from the form
            Dim fields As System.Collections.Generic.IDictionary(Of String, iTextSharp.text.pdf.AcroFields.Item) = frm.Fields
            Dim columns As DataColumnCollection = fieldData.Table.Columns
            For Each key As String In fields.Keys
                If columns.Contains(key) Then
                    frm.SetField(key, fieldData.Item(key).ToString)
                End If
            Next
            stamper.Close()
            fs.Close()
            reader.Close()
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' A demo for filling a pdf form. Dynamically adding more pages to fit long text for a field.
    ''' </summary>
    ''' <param name="sourcePdf"></param>
    ''' <param name="fieldData"></param>
    ''' <param name="outputPdf"></param>
    ''' <remarks>Demo filling the History field with very long text</remarks>
    Public Shared Sub FillMyForm(ByVal sourcePdf As String, ByVal fieldData As DataRow, ByVal outputPdf As String)
        Try
            'Open the pdf using pdfreader
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePdf)

            'create a filestream for output
            Dim fs As New System.IO.FileStream(outputPdf, IO.FileMode.Create, IO.FileAccess.Write)

            'use stamper to copy the source pdf to output
            Dim stamper As New iTextSharp.text.pdf.PdfStamper(reader, fs)

            'Get the form from the pdf
            Dim frm As iTextSharp.text.pdf.AcroFields = stamper.AcroFields

            'get the fields from the form
            Dim fields As System.Collections.Generic.IDictionary(Of String, iTextSharp.text.pdf.AcroFields.Item) = frm.Fields

            'get the size of the history textfield in the pdf in term of a rectangle. This will be used
            'later for calculating how much text would fit in it.
            Dim boxSize As IList(Of iTextSharp.text.pdf.AcroFields.FieldPosition) = frm.GetFieldPositions("History")
            Dim rect As iTextSharp.text.Rectangle = boxSize(0).position
            'By manually typing in the history field, I found that it can hold 33 lines
            Dim boxMaxLines As Integer = 32

            'create a basefont which is used later to measure the text. This needs to match the font used in the pdf field. 
            Dim fnt As iTextSharp.text.pdf.BaseFont = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.TIMES_ROMAN, "Cp1252", False)

            'Get the width of 1 text line
            Dim aLine As String = "This text will fit in a single line of history field using times-roman font @ size 8.0. This text will fit in a single line of history field using times-ro"
            'This is the maximum number of characters that would fit in the history field
            Dim charLimit As Integer = aLine.Length * (boxMaxLines - 1)

            'Now loop thru the items in fieldData datarow and set the value to the pdf fields.
            'Note that we use the datatable column names as the pdf field names. So when you create your
            'datatable, since pdf field names are case-sensitive, you need to name the columns exactly the 
            'same as your pdf field names.
            For Each col As DataColumn In fieldData.Table.Columns
                'Get the name of this column, which is also the pdf field name
                Dim key As String = col.ColumnName

                'If we're filling the history field, we need to measure the text and split it
                'onto the next page if it's too long.
                If key.ToUpper = "HISTORY" Then
                    Dim histTxt As String = fieldData.Item(key).ToString()
                    If histTxt.Length > charLimit Then
                        'We need to split the text into 2 parts
                        Dim part1 As String = histTxt.Substring(0, charLimit - aLine.Length)
                        part1 = part1.Substring(0, part1.LastIndexOf(" "c))
                        Dim part2 As String = histTxt.Replace(part1, "")

                        'Set the field value using part1
                        frm.SetField(key, part1)

                        'We then add a new page to put part2 in
                        Dim newPageNumber As Integer = reader.NumberOfPages + 1
                        stamper.InsertPage(newPageNumber, reader.GetPageSize(1))
                        Dim newRect As New iTextSharp.text.Rectangle(rect.Left, 750, rect.Right, 50)
                        Dim txtField As New iTextSharp.text.pdf.TextField(stamper.Writer, newRect, "HistoryCont")
                        With txtField
                            'Set the properties of this text field to match the other one.
                            .BorderStyle = iTextSharp.text.pdf.PdfBorderDictionary.STYLE_SOLID
                            .Font = fnt
                            .FontSize = 8.0
                            .BorderColor = iTextSharp.text.BaseColor.BLACK
                            .BorderWidth = 0.5
                            .Options = iTextSharp.text.pdf.TextField.MULTILINE
                            .Text = part2
                        End With
                        stamper.AddAnnotation(txtField.GetTextField, newPageNumber)
                    Else
                        frm.SetField(key, histTxt)
                    End If
                Else
                    'With all other fields, we just go ahead and set the field value
                    frm.SetField(key, fieldData.Item(key).ToString())
                End If
            Next
            'After done, do the cleanup
            stamper.Close()
            fs.Close()
            reader.Close()
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try
    End Sub


    Public Shared Sub AddTextAnnotation(ByVal sourcePdf As String, ByVal outputPdf As String)

        Try
            'Open the pdf using pdfreader
            Dim reader As New iTextSharp.text.pdf.PdfReader(sourcePdf)
            'create a filestream for output
            Dim fs As New System.IO.FileStream(outputPdf, IO.FileMode.Create, IO.FileAccess.Write)
            'use stamper to copy the source pdf to output
            Dim stamper As New iTextSharp.text.pdf.PdfStamper(reader, fs)
            'Create an annotation
            Dim rect As New iTextSharp.text.Rectangle(100, 500, 120, 520) 'The rectangle that represents the annotation
            Dim title As String = "Comments"
            Dim annotText As String = "This is some text that will be displayed as an annotation."
            Dim shouldOpen As Boolean = True
            Dim iconStyle As String = "Comment"
            Dim pageNumber As Integer = 1 'The page number to add this annotation to
            Dim annot As iTextSharp.text.pdf.PdfAnnotation = iTextSharp.text.pdf.PdfAnnotation.CreateText(stamper.Writer, rect, "Comments", annotText, True, "Comments")
            stamper.AddAnnotation(annot, pageNumber)
            stamper.Close()
            reader.Close()
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try
    End Sub

    Public Shared Function GetAcroFieldsData(ByVal sourcePdf As String) As Dictionary(Of String, String)
        Dim frmData As New Dictionary(Of String, String)
        Try
            'Open the pdf using pdfreader
            Dim reader As New PdfReader(sourcePdf)
            'Get the form from the pdf
            Dim frm As AcroFields = reader.AcroFields
            'get the fields from the form

            Dim fields As System.Collections.Generic.IDictionary(Of String, iTextSharp.text.pdf.AcroFields.Item) = frm.Fields
            'Extract the data from the fields
            Dim data As String = String.Empty
            For Each key As String In fields.Keys
                data = frm.GetField(key)
                frmData.Add(key, data)
            Next
            reader.Close()
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try
        Return frmData
    End Function

    Public Shared Function GetPdfSummary(ByVal sourcePdf As String) As DataTable
        Dim summary As New DataTable()
        With summary.Columns
            .Add("FileName", GetType(String))
            .Add("PageCount", GetType(Integer))
            .Add("PageSize", GetType(String))
            .Add("Title", GetType(String))
            .Add("Author", GetType(String))
            .Add("Subject", GetType(String))
        End With
        Try
            'Open the pdf using pdfreader and get summary data
            Dim raf As iTextSharp.text.pdf.RandomAccessFileOrArray = New iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf)
            Dim reader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(raf, Nothing)
            Dim row As DataRow = summary.NewRow()
            row("FileName") = System.IO.Path.GetFileName(sourcePdf)
            row("PageCount") = reader.NumberOfPages()
            Dim rect As iTextSharp.text.Rectangle = reader.GetPageSize(1)
            row("PageSize") = String.Format("{0} x {1} in", (rect.Width / 72).ToString("f2"), (rect.Height / 72).ToString("f2"))
            Dim info As System.Collections.Generic.Dictionary(Of String, String) = reader.Info
            For Each key As String In info.Keys
                Select Case key
                    Case "Title"
                        row("Title") = info.Item(key)
                    Case "Author"
                        row("Author") = info.Item(key)
                    Case "Subject"
                        row("Subject") = info.Item(key)
                End Select
            Next
            summary.Rows.Add(row)
            reader.Close()
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try
        Return summary
    End Function


    ''' <summary>
    ''' Replace specified page(s) in a pdf file with a blank page
    ''' </summary>
    ''' <param name="sourcePdf">The source pdf file to have pages replaced</param>
    ''' <param name="pagesToReplace">List of pages in source pdf to be replaced with blank page</param>
    ''' <param name="outPdf">The output pdf with pages replaced by blank pages</param>
    ''' <param name="templatePdf">Optional template pdf to used as replacement page</param>
    ''' <returns>True if successful, False if failed</returns>
    ''' <remarks>If the template pdf parameter is left blank, a blank template pdf will be created on the fly
    ''' and deleted when done</remarks>
    Public Shared Function ReplacePagesWithBlank(ByVal sourcePdf As String, _
                                                 ByVal pagesToReplace As List(Of Integer), _
                                                 ByVal outPdf As String, _
                                                 Optional ByVal templatePdf As String = "") As Boolean
        Dim result As Boolean = False
        Dim template As iTextSharp.text.pdf.PdfReader = Nothing
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim copier As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim deleteTemplate As Boolean = False
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            If String.IsNullOrEmpty(templatePdf) Then
                templatePdf = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\BlankPageTemplate.pdf"
                deleteTemplate = True
                Dim tpdoc As New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
                Dim tpwriter As iTextSharp.text.pdf.PdfWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(tpdoc, New IO.FileStream(templatePdf, IO.FileMode.Create))
                tpdoc.Open()
                tpdoc.Add(New iTextSharp.text.Paragraph(" "))
                tpdoc.Close()
            End If
            template = New iTextSharp.text.pdf.PdfReader(templatePdf)
            doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
            copier = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outPdf, IO.FileMode.Create))
            doc.Open()
            For i As Integer = 1 To reader.NumberOfPages
                If pagesToReplace.Contains(i) Then
                    copier.AddPage(copier.GetImportedPage(template, 1))
                Else
                    copier.AddPage(copier.GetImportedPage(reader, i))
                End If
            Next
            doc.Close()
            template.Close()
            reader.Close()
            result = True
            If deleteTemplate Then
                IO.File.Delete(templatePdf)
            End If
        Catch ex As Exception
            'Put your own code to handle exception here
            Debug.Write(ex.Message)
        End Try
        Return result
    End Function

    ''' <summary>
    ''' Insert new pages to an existing pdf file
    ''' </summary>
    ''' <param name="sourcePdf">The full path to the source pdf</param>
    ''' <param name="pagesToInsert">The dictionary contains the pages to be inserted in the source pdf. The key is the page number to be inserted. The value is the PdfImportedPage to insert</param>
    ''' <param name="outPdf">The full path of the resulting output pdf file</param>
    ''' <returns>True if the operation succeeded. False otherwise.</returns>
    ''' <remarks>To create the pagesToInsert dictionary, you can use the iTextSharp.text.pdf.PdfCopy class to open
    ''' an existing pdf file and call the GetImportedPage method</remarks>
    Public Shared Function InsertPages(ByVal sourcePdf As String, _
                                       ByVal pagesToInsert As Dictionary(Of Integer, iTextSharp.text.pdf.PdfImportedPage), _
                                       ByVal outPdf As String) As Boolean
        Dim result As Boolean = False
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim copier As iTextSharp.text.pdf.PdfCopy = Nothing
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
            copier = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outPdf, IO.FileMode.Create))
            doc.Open()
            For i As Integer = 1 To reader.NumberOfPages
                If pagesToInsert.ContainsKey(i) Then
                    copier.AddPage(pagesToInsert(i))
                End If
                copier.AddPage(copier.GetImportedPage(reader, i))
            Next
            doc.Close()
            reader.Close()
            result = True
        Catch ex As Exception
            'Put your own code to handle exception here
            Debug.Write(ex.Message)
        End Try
        Return result
    End Function


    ''' <summary>
    ''' Remove specified page(s) from a pdf file
    ''' </summary>
    ''' <param name="sourcePdf">The source pdf to have pages removed from</param>
    ''' <param name="pagesToRemove">List of pages to be removed</param>
    ''' <param name="outputPdf">The output pdf after pages removed</param>
    ''' <returns>True if successful, false if failed</returns>
    ''' <remarks></remarks>
    Public Shared Function RemovePages(ByVal sourcePdf As String, ByVal pagesToRemove As List(Of Integer), ByVal outputPdf As String) As Boolean
        Dim result As Boolean = False
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim copier As iTextSharp.text.pdf.PdfCopy = Nothing
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
            copier = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outputPdf, IO.FileMode.Create))
            doc.Open()
            For i As Integer = 1 To reader.NumberOfPages
                If Not pagesToRemove.Contains(i) Then
                    copier.AddPage(copier.GetImportedPage(reader, i))
                End If
            Next
            doc.Close()
            reader.Close()
            result = True
        Catch ex As Exception
            'Put your own code to handle exception here
            Debug.Write(ex.Message)
        End Try
        Return result
    End Function

    Public Overloads Shared Sub CreateBlankPdf(ByVal pageSize As iTextSharp.text.Rectangle, ByVal outPdf As String)
        Dim doc As iTextSharp.text.Document = Nothing
        Dim writer As iTextSharp.text.pdf.PdfWriter = Nothing
        Try
            doc = New iTextSharp.text.Document(pageSize)
            writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, New IO.FileStream(outPdf, IO.FileMode.Create))
            doc.Open()
            doc.Add(New iTextSharp.text.Paragraph(" "))
            doc.Close()
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try
    End Sub

    Public Overloads Shared Sub CreateBlankPdf(ByVal sourcePdf As String, ByVal outPdf As String)
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim writer As iTextSharp.text.pdf.PdfWriter = Nothing
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            doc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
            writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, New IO.FileStream(outPdf, IO.FileMode.Create))
            doc.Open()
            doc.Add(New iTextSharp.text.Paragraph(" "))
            doc.Close()
        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try
    End Sub

    Public Shared Sub DrawShapesDemo(ByVal sourcePdf As String, ByVal outputPdf As String)
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim stamper As iTextSharp.text.pdf.PdfStamper = Nothing
        Dim cb As iTextSharp.text.pdf.PdfContentByte = Nothing
        Dim rect As iTextSharp.text.Rectangle = Nothing
        Dim pageCount As Integer = 0
        Dim borderColor, fillColor As iTextSharp.text.BaseColor
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            rect = reader.GetPageSizeWithRotation(1)
            stamper = New iTextSharp.text.pdf.PdfStamper(reader, New System.IO.FileStream(outputPdf, IO.FileMode.Create))
            ' Set up the border and fill color for the shapes to be drawn
            borderColor = iTextSharp.text.BaseColor.BLUE
            fillColor = iTextSharp.text.BaseColor.PINK
            'Loop thru the pages and draw various shapes to their overcontent layer
            pageCount = reader.NumberOfPages()
            For i As Integer = 1 To pageCount
                'Get the undercontent layer of this page
                cb = stamper.GetUnderContent(i)
                '<<< Note: if you want the drawings to appear on top of the contents (covering it)
                ' then you need to get the overcontent layer.
                'cb = stamper.GetOverContent(i)

                'Set the boder color of the shapes
                cb.SetColorStroke(borderColor)
                'Set the fill color of the shapes
                cb.SetColorFill(fillColor)
                'Start drawing shapes. 

                ' >>>> Remember, the cordinate of the LOWER-LEFT corner of a page is (0, 0)
                ' 1 in = 72 units, so a 8.5 x 11 page will have a width of 612 units and a height of 792 units.
                ' Figuring out where to draw your shapes will be much easier if you use a piece of paper to
                ' plot out the cordinates first.

                'Draw a circle centered at (135, 500) with a radius of 50
                cb.Circle(135, 500, 50)
                'Draw an ellipse that fits in a ractangle with (190, 450) as the lower-left corner
                'and (400, 550) as the upper-right corner
                cb.Ellipse(190, 450, 400, 550)
                'Draw a square with the lower-left corner is (410, 450) and the width (and height) = 100
                cb.Rectangle(410, 450, 100, 100)
                'Draw a rounded rectangle
                cb.RoundRectangle(150, 330, 200, 100, 20)
                'Color fill the shapes above
                cb.FillStroke()
                'Draw a line starting from (150, 310) to (450, 310)
                cb.MoveTo(150, 310)
                cb.LineTo(450, 310)
                cb.Stroke()
                'Draw a triangle with vertices (290, 300), (150, 150) and (450, 150) without filling
                cb.MoveTo(290, 300)
                cb.LineTo(150, 150)
                cb.LineTo(450, 150)
                cb.ClosePathStroke()
            Next
            stamper.Close()
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Add an image to pdf pages
    ''' </summary>
    ''' <param name="sourcePdf">The full path to the source pdf file</param>
    ''' <param name="outputPdf">The full path to the image file to use</param>
    ''' <param name="imgPath">The full path to be used for the output pdf file</param>
    ''' <param name="imgLocation">The lower left corner location where the image will be placed on a pdf page</param>
    ''' <param name="imgSize">The size of the image on pdf page</param>
    ''' <param name="pages">The pages where the image should be added. Default option is Nothing which will add the image to all pages on the pdf file</param>
    ''' <remarks>Units for location and size are in points. 1 inch = 72 points</remarks>
    Public Shared Sub AddImageToPage(ByVal sourcePdf As String, ByVal outputPdf As String, ByVal imgPath As String, ByVal imgLocation As Point, ByVal imgSize As Size, Optional ByVal pages() As Integer = Nothing)
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim stamper As iTextSharp.text.pdf.PdfStamper = Nothing
        Dim img As iTextSharp.text.Image = Nothing
        Dim cb As iTextSharp.text.pdf.PdfContentByte = Nothing
        Dim pageCount As Integer = 0
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            stamper = New iTextSharp.text.pdf.PdfStamper(reader, New System.IO.FileStream(outputPdf, IO.FileMode.Create))
            img = iTextSharp.text.Image.GetInstance(imgPath)
            img.ScaleAbsolute(imgSize.Width, imgSize.Height)
            img.SetAbsolutePosition(imgLocation.X, imgLocation.Y)
            pageCount = reader.NumberOfPages()
            If pages IsNot Nothing Then
                For Each page As Integer In pages
                    If page > 0 AndAlso page <= pageCount Then
                        cb = stamper.GetOverContent(page)
                        cb.AddImage(img)
                    End If
                Next
            Else
                For i As Integer = 1 To pageCount
                    cb = stamper.GetOverContent(i)
                    cb.AddImage(img)
                Next
            End If
            stamper.Close()
            reader.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Extract specfied pages from a pdf and insert them to another pdf.
    ''' </summary>
    ''' <param name="pdfToExtract">the full path to the source pdf whose the pages are to be extracted from</param>
    ''' <param name="pdfToInsert">the full path to the source pdf whose extracted pages are to be inserted into</param>
    ''' <param name="insertionDetails">a dictionary whose key is the page number to be extracted and value is the page number to be inserted</param>
    ''' <param name="outputPdf">full path to the resulting pdf</param>
    ''' <returns>True if successful. Otherwise false</returns>
    ''' <remarks>insertionDetails dictionary example. Let's say you want to extract page # 2 from pdfToExtract file
    ''' and insert that page as page # 5 in pdfToInsert file. So you will create a dictionary to pass to this function and add
    ''' an entery to it like this:
    ''' Dim myDict as New Dictionary(Of Integer, Integer)
    ''' myDict.Add(2, 5)
    ''' </remarks>
    Public Function ExtractAndInsertPages(ByVal pdfToExtract As String, _
                                          ByVal pdfToInsert As String, _
                                          ByVal insertionDetails As Dictionary(Of Integer, Integer), _
                                          ByVal outputPdf As String) As Boolean
        Dim result As Boolean = False
        Dim reader1 As iTextSharp.text.pdf.PdfReader = Nothing
        Dim reader2 As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim copier As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Dim pagesToInsert As New Dictionary(Of Integer, iTextSharp.text.pdf.PdfImportedPage)
        Try
            reader1 = New iTextSharp.text.pdf.PdfReader(pdfToExtract)
            reader2 = New iTextSharp.text.pdf.PdfReader(pdfToInsert)
            doc = New iTextSharp.text.Document(reader2.GetPageSizeWithRotation(1))
            copier = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outputPdf, IO.FileMode.Create))
            doc.Open()
            'First extract the pages from pdfToExtract file and put in a dictionary
            For Each pageNum As Integer In insertionDetails.Keys
                pagesToInsert.Add(insertionDetails(pageNum), copier.GetImportedPage(reader1, pageNum))
            Next
            reader1.Close()

            'Now insert the pages from dictionary to pdfToInsert file
            For i As Integer = 1 To reader2.NumberOfPages
                If pagesToInsert.ContainsKey(i) Then
                    copier.AddPage(pagesToInsert(i))
                End If
                copier.AddPage(copier.GetImportedPage(reader2, i))
            Next
            doc.Close()
            reader2.Close()
            result = True
        Catch ex As Exception
            'Put your own code to handle exception here
            Debug.Write(ex.Message)
        End Try
        Return result
    End Function

    Public Overloads Shared Function ResizePage(ByVal sourcePdf As String, ByVal newSize As iTextSharp.text.Rectangle, ByVal outputPdf As String) As Boolean
        Dim result As Boolean = False
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim copier As iTextSharp.text.pdf.PdfCopy = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim pageDict As iTextSharp.text.pdf.PdfDictionary = Nothing
        Dim pageCount As Integer = 0
        Dim rect As iTextSharp.text.pdf.PdfRectangle = Nothing
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            doc = New iTextSharp.text.Document(newSize)
            copier = New iTextSharp.text.pdf.PdfCopy(doc, New IO.FileStream(outputPdf, IO.FileMode.Create))
            pageCount = reader.NumberOfPages()
            rect = New iTextSharp.text.pdf.PdfRectangle(newSize)
            doc.Open()
            For i As Integer = 1 To pageCount
                pageDict = reader.GetPageN(i)
                pageDict.Put(PdfName.MEDIABOX, rect)
                pageDict.Put(PdfName.CROPBOX, rect)
                copier.AddPage(copier.GetImportedPage(reader, i))
            Next
            doc.Close()
            reader.Close()
            result = True
        Catch ex As Exception
            'Put your own code to handle any exception here
        End Try
        Return result
    End Function

    Public Overloads Shared Function ResizePage2(ByVal sourcePdf As String, ByVal newSize As iTextSharp.text.Rectangle, ByVal outputPdf As String) As Boolean
        Dim result As Boolean = False
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim doc As iTextSharp.text.Document = Nothing
        Dim writer As iTextSharp.text.pdf.PdfWriter = Nothing
        Dim pageCount As Integer = 0
        Dim cb As iTextSharp.text.pdf.PdfContentByte = Nothing
        Dim page As iTextSharp.text.pdf.PdfImportedPage = Nothing
        Dim x, y As Single
        Try
            reader = New iTextSharp.text.pdf.PdfReader(sourcePdf)
            doc = New iTextSharp.text.Document(newSize)
            writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, New System.IO.FileStream(outputPdf, IO.FileMode.Create))
            pageCount = reader.NumberOfPages()
            doc.Open()
            cb = writer.DirectContent()
            For i As Integer = 1 To pageCount
                doc.NewPage()
                page = writer.GetImportedPage(reader, i)
                x = (newSize.Width - reader.GetCropBox(i).Width) / 2
                y = (newSize.Height - reader.GetCropBox(i).Height) / 2
                cb.AddTemplate(page, x, y)
            Next
            doc.Close()
            reader.Close()
            result = True
        Catch ex As Exception
            'Put your own code to handle any exception here
        End Try
        Return result
    End Function

End Class
