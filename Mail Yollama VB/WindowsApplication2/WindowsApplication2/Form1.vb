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
        TextBox6.Clear()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim gonderilecek, gonderen, mailkonusu, iletilecekmail As String
            gonderilecek = TextBox3.Text
            gonderen = TextBox1.Text
            mailkonusu = TextBox4.Text
            iletilecekmail = TextBox5.Text
            Dim gonderici As New MailAddress(gonderen)
            Dim alici As New MailAddress(gonderilecek)
            Dim eposta As New MailMessage(gonderici, alici)
            eposta.IsBodyHtml = True
            eposta.Subject = mailkonusu
            Dim mailbilgileri As New NetworkCredential(TextBox1.Text, TextBox2.Text)
            eposta.Body = iletilecekmail
            Dim mailclient As New SmtpClient
            mailclient.Host = "smtp.gmail.com"
            mailclient.Port = "587"
            mailclient.EnableSsl = True
            mailclient.Credentials = mailbilgileri
            mailclient.Send(eposta)
            MsgBox("mesajýnýz baþarý ile gönderilmiþtir")

        Catch
            MsgBox("beklenmeyen hata ")
        End Try


    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Dim gonderilecek, gonderen, mailkonusu, iletilecekmail As String
            gonderilecek = TextBox3.Text
            gonderen = TextBox1.Text
            mailkonusu = TextBox4.Text
            iletilecekmail = TextBox5.Text
            Dim gonderici As New MailAddress(gonderen, mailkonusu)
            Dim alici As New MailAddress(gonderilecek)
            Dim eposta As New MailMessage(gonderici, alici)
            eposta.IsBodyHtml = True
            eposta.Subject = mailkonusu
            Dim mailbilgileri As New NetworkCredential(TextBox1.Text, TextBox2.Text)
            eposta.Body = iletilecekmail
            eposta.AlternateViews.Add(New AlternateView(yol))
            Dim mailclient As New SmtpClient
            mailclient.Host = "smtp.gmail.com"
            mailclient.Port = "587"
            mailclient.EnableSsl = True
            mailclient.Credentials = mailbilgileri
            mailclient.Send(eposta)
            MsgBox("mesajýnýz baþarý ile gönderilmiþtir")

        Catch
            MsgBox("beklenmeyen hata ")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            yol = OpenFileDialog1.FileName
            TextBox6.Text = yol


        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim sayi As Integer = TextBox7.Text
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = sayi
        Dim i As Integer
        For i = 1 To sayi



            Dim gonderilecek, gonderen, mailkonusu, iletilecekmail As String
            gonderilecek = TextBox3.Text
            gonderen = TextBox1.Text
            mailkonusu = TextBox4.Text
            iletilecekmail = TextBox5.Text
            Dim gonderici As New MailAddress(gonderen, mailkonusu)
            Dim alici As New MailAddress(gonderilecek)
            Dim eposta As New MailMessage(gonderici, alici)
            eposta.IsBodyHtml = True
            eposta.Subject = mailkonusu
            Dim mailbilgileri As New NetworkCredential(TextBox1.Text, TextBox2.Text)
            eposta.Body = iletilecekmail
            Dim mailclient As New SmtpClient
            mailclient.Host = "smtp.gmail.com"
            mailclient.Port = "587"
            mailclient.EnableSsl = True
            mailclient.Credentials = mailbilgileri
            mailclient.Send(eposta)
            ProgressBar1.Value += True



        Next
        MsgBox("mesajýnýz baþarý ile gönderilmiþtir !!")


    End Sub
End Class
