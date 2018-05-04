' handy dandy code for reading in files/using file explorer to choose files etc...

Imports System.Text.RegularExpressions

Public Class VB_CH9_Sup

    Private strWorkingString As String

    Private Sub LoadToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSelectFile.Click
        Using ofd As New OpenFileDialog

            ofd.Filter = "txt files(*.txt)|*.txt|All files (*.*)|*.*"

            If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then

                txtFilePath.Text = ofd.FileName
            End If
        End Using
    End Sub

    Private Sub btnProcess_Click(sender As Object, e As EventArgs) Handles btnProcess.Click
        If (txtFilePath.Text) = "" Then
            MsgBox("You must select a text file!")
        Else

            strWorkingString = IO.File.ReadAllText(txtFilePath.Text)

            Dim extension As String = System.IO.Path.GetExtension(txtFilePath.ToString)

            If (extension = ".xlsx") Or (extension = ".xls") Then

            Else
                MsgBox("Please choose a .xlsx file or a .xls file", 0, "Wrong file type")
            End If
            txtDisplayText.Text = strWorkingString

            TabControl1.SelectTab(1)
        End If
    End Sub

    Private Sub btnRemPunct_Click(sender As Object, e As EventArgs) Handles btnRemPunct.Click

        Dim strNoPunctuation As String

        strNoPunctuation = Regex.Replace(strWorkingString, "[^A-Za-z0-9 \r\n]+", String.Empty)

        txtDisplayText.Text = strNoPunctuation
    End Sub

    Private Sub btnMakeLower_Click(sender As Object, e As EventArgs) Handles btnMakeLower.Click
        Dim strLowerCase As String

        strLowerCase = txtDisplayText.Text.ToLower

        txtDisplayText.Text = strLowerCase

    End Sub

    ' a word counter function
    Private Sub btnCountToFile_Click(sender As Object, e As EventArgs) Handles btnCountToFile.Click
        Dim WordCountWriter As IO.StreamWriter
        WordCountWriter = My.Computer.FileSystem.OpenTextFileWriter("wordCount.txt", False)
        Dim intWordCount As Integer = 0
        Dim intTotalWordCount As Integer = 0

        Dim Words As New Dictionary(Of String, Integer)
        Dim intKeyValue As Integer

        Dim strKey As String
        Dim strValue As String

        Dim strCountWords() As String = Regex.Split(strWorkingString, "\W+")
        Dim strSingleWord As String

        For Each strSingleWord In strCountWords
            If (Words.ContainsKey(strSingleWord)) Then
                intKeyValue = Words(strSingleWord)
                intKeyValue += 1
                Words(strSingleWord) = intKeyValue
            Else Words.Add(strSingleWord, 1)
                intWordCount += 1
            End If
        Next

        Dim WordPair As KeyValuePair(Of String, Integer)

        For Each WordPair In Words
            strKey = Convert.ToString(WordPair.Key)
            strValue = Convert.ToString(WordPair.Value)
            intTotalWordCount = intTotalWordCount + WordPair.Value
            WordCountWriter.WriteLine(strKey + "," + strValue)
        Next

        WordCountWriter.WriteLine("word count: " & intTotalWordCount)
        WordCountWriter.Close()
        MessageBox.Show("write complete", "writefile", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub close_Click(sender As Object, e As EventArgs) Handles close.Click
        Application.Exit()
    End Sub
End Class
