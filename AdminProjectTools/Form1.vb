Imports System.Globalization

Public Class Form1
    Private selisih As Single = 0.0F
    Private formattedKlaim As String = "Rp 0"
    Private totalKlaim As Single = 0.0F

    Dim machineCC As New Dictionary(Of String, Integer)() From {
        {">150", 15},
        {"<150", 20}
    }

    Dim more As Integer = machineCC(">150")
    Dim less As Integer = machineCC("<150")
    Dim modifierCc As Integer


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    Private Sub Calculate()
        cleanVar()

        modifierCc = less
        Try
            If Not (TextBox1.Text = "" Or TextBox2.Text = "") Then
                selisih = Single.Parse(TextBox2.Text) - Single.Parse(TextBox1.Text)
                If (RadioButton1.Checked = True) Then
                    modifierCc = less
                    totalKlaim = (selisih / modifierCc) * 10000.0F
                ElseIf (RadioButton2.Checked = True) Then
                    modifierCc = more
                    totalKlaim = (selisih / modifierCc) * 10000.0F
                End If

                formattedKlaim = "Rp " & totalKlaim.ToString("#,##0.00")
                'formattedKlaim = "Rp " & totalKlaim.ToString("F0")

                Dim ccVal As Integer = 0
                TextBox3.AppendText(String.Format("Gasoline Claim Rule: {5}1. <150cc = difference(in km)/20 * 10.000{5}2. >150cc = difference(in Km)/15 * 10.000{5}{5}Known Refference --> Difference {0}Km,{5}{5}Claim Gasoline = ({0}/{1})*10000{5}Claim Gasoline({2}-{3} Km) = {4}{5}",
                    selisih, modifierCc, TextBox1.Text, TextBox2.Text, formattedKlaim, vbCrLf))

            End If



        Catch ex As Exception
            TextBox3.AppendText("Err: " & ex.Message)
        End Try
    End Sub

    Private Sub cleanVar()
        selisih = 0
        totalKlaim = 0
        TextBox3.Clear()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Calculate()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Calculate()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Calculate()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Calculate()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If Not (TextBox1.Text = vbNull Or TextBox2.Text = vbNull) Then
                Dim lineCount As Integer = TextBox4.Lines.Length
                If lineCount > 0 AndAlso Not String.IsNullOrWhiteSpace(TextBox4.Lines(lineCount - 1)) Then
                    TextBox4.AppendText(vbCrLf)
                End If

                TextBox4.AppendText(String.Format("Claim Gasoline({0}-{1} Km): {2}{3}",
                    TextBox1.Text, TextBox2.Text, formattedKlaim, vbCrLf))
                totaling()
            End If
        Catch ex As Exception
        End Try
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox4.Lines.Length > 0 Then
            TextBox4.Text = String.Join(vbCrLf, TextBox4.Lines.Take(TextBox4.Lines.Length - 1))
            totaling()
        End If
    End Sub


    Private Sub totaling()
        TextBox5.Clear()
        Dim total As Decimal = 0.0D

        For Each line As String In TextBox4.Lines
            If line.StartsWith("Claim Gasoline") AndAlso line.Contains("Rp") Then
                Dim startIndex As Integer = line.IndexOf("Rp") + 3
                Dim endIndex As Integer = line.IndexOf(".", startIndex)

                Dim amountString As String
                If endIndex > 0 Then
                    amountString = line.Substring(startIndex, endIndex - startIndex + 3)
                Else
                    amountString = line.Substring(startIndex)
                End If

                Dim amount As Decimal = Decimal.Parse(amountString, CultureInfo.InvariantCulture)
                total += amount
            End If
        Next


        TextBox5.Text = "Total: Rp " & total.ToString("#,##0.00")
    End Sub




    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub
End Class
