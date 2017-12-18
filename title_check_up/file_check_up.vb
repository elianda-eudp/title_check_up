Imports Microsoft.Office.Interop.Word

Module file_check_up
    Private comment_str As String

    Sub check_title_up() '标题下无正文检查

        Dim ShapesNum As Long, i As Long, j As Long
        'Application.ScreenUpdating = False
        Dim t As Word.Paragraph
        Dim flag As Integer
        flag = 0
        Dim wd As Word.Document
        wd = Globals.ThisAddIn.Application.ActiveDocument
        For i = 1 To wd.Paragraphs.Count

            If i > wd.Paragraphs.Count Then
                Exit For
            End If
            t = wd.Paragraphs(i)
            t.Range.Select()


            If flag = 0 And wd.ActiveWindow.Selection.Information(WdInformation.wdWithInTable) Then
                Dim col As Integer
                col = wd.ActiveWindow.Selection.Information(Word.WdInformation.wdEndOfRangeColumnNumber)
                Dim Row As Integer
                Row = wd.ActiveWindow.Selection.Information(Word.WdInformation.wdEndOfRangeRowNumber)
                i = i + Row * (col + 1) - 2
                flag = 1
            ElseIf flag = 1 And wd.ActiveWindow.Selection.Information(Word.WdInformation.wdWithInTable) Then

            Else
                flag = 0
                If Len(Trim(wd.Paragraphs(i).Range.Text)) = 1 Then
                    wd.Paragraphs(i).Range.Delete
                    i = i - 1
                Else
                    If i + 1 <= wd.Paragraphs.Count Then
                        If t.Format.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText Or wd.Paragraphs(i + 1).Format.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText Then

                        ElseIf t.Format.OutlineLevel >= wd.Paragraphs(i + 1).Format.OutlineLevel And wd.Paragraphs(i + 1).Format.OutlineLevel < Word.WdOutlineLevel.wdOutlineLevelBodyText Then
                            Dim err_str As String
                            Dim return_val As Boolean
                            Dim comments As Word.Comments
                            Dim comment As Word.Comment
                            comment_str = "标题下无正文，内容错误!" & vbCrLf

                            comments = wd.comments
                            comment = comments.Add(t.Range, comment_str)
                            comment.Author = "coin wo-wo"
                            comment.Range.Text = comment_str
                        End If
                    End If
                End If
            End If
        Next i
        wd.ScreenUpdating = True

    End Sub



End Module
