Imports System.Net
Imports System.Net.Mail
Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CheckedListBox1.Items.Add(TextBox3.Text & vbTab & TextBox6.Text & vbTab & TextBox7.Text)
        TextBox3.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox3.Select()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim i, sayi As Integer
            sayi = CheckedListBox1.Items.Count
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = sayi

            For i = 0 To sayi - 1
                Dim kisi As String = CheckedListBox1.Items(i)
                Dim dizi() As String = kisi.Split("	")
                Dim kime, kimden, konu, sifre, mesaj As String
                kime = CheckedListBox1.SelectedItem
                kime = dizi(2)
                kimden = TextBox1.Text
                sifre = TextBox2.Text
                konu = TextBox4.Text
                mesaj = TextBox5.Text

                Dim gonderen As New MailAddress(kimden)
                Dim alici As New MailAddress(kime)
                Dim eposta As New MailMessage(kimden, kime)
                eposta.IsBodyHtml = True
                eposta.Subject = konu
                eposta.Body = mesaj

                Dim mailistemcisi As New Mail.SmtpClient
                Dim kimlik As New NetworkCredential(kimden, sifre)
                mailistemcisi.Host = "smtp.gmail.com"
                mailistemcisi.Port = "587"
                mailistemcisi.EnableSsl = True
                mailistemcisi.UseDefaultCredentials = True
                mailistemcisi.Credentials = kimlik
                mailistemcisi.Send(eposta)
                ProgressBar1.Value += 1
                Label7.Text = "G�nderiliyor..." & vbTab & i & " / " & sayi
            Next
            MsgBox("<<<MA�L G�NDER�LD�>>>>!!!!")
        Catch
            MsgBox("Beklenmeyen Bir  Hata Olu�tu ",MsgBoxStyle.Critical)

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CheckedListBox1.Items.Clear()


    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox4.Clear()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TextBox5.Clear()

    End Sub
End Class
