Imports System.Net
Imports System.Net.Mail
Public Class Form1
    Dim yol As String
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim kime, kimden, konu, sifre, mesaj As String
        kime = TextBox3.Text
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
        MsgBox("Mesajýnýz iletildi!!!")

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            yol = OpenFileDialog1.FileName
            TextBox6.Text = yol

        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim kime, kimden, konu, sifre, mesaj As String
        kime = TextBox3.Text
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
        eposta.Attachments.Add(New Attachment(yol))
        Dim mailistemcisi As New Mail.SmtpClient
        Dim kimlik As New NetworkCredential(kimden, sifre)
        mailistemcisi.Host = "smtp.gmail.com"
        mailistemcisi.Port = "587"
        mailistemcisi.EnableSsl = True
        mailistemcisi.UseDefaultCredentials = True
        mailistemcisi.Credentials = kimlik
        mailistemcisi.Send(eposta)
        MsgBox("Mesajýnýz iletildi!!!")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = TextBox7.Text
        Dim i, sayi As Integer
        sayi = TextBox7.Text
        For i = 1 To sayi

            Dim kime, kimden, konu, sifre, mesaj As String
            kime = TextBox3.Text
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
        Next

        MsgBox("Mesajýnýz iletildi!!!")

    End Sub

   

   
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Hide()
        Form2.Show()

    End Sub
End Class
