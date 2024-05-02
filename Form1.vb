Imports System.IO
Imports System.ComponentModel
Imports System.Net.Mail
Public Class Form1

    Dim Seconds As Integer = 0
    Dim DailyPath As String = "D:\Data\www\VacationCountDownEmailer\Data\Daily.htm"
    Dim TheAppDataPath As String = "D:\Data\www\VacationCountDownEmailer\Data\"


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Show()
        Application.DoEvents()
        System.Threading.Thread.Sleep(30000) 'give the system time to wake up if this is being started by a service...

        Dim strLastModified = System.IO.File.GetLastWriteTime(DailyPath).ToShortDateString()

        'If strLastModified = Now.ToShortDateString Then
        '    System.Threading.Thread.Sleep(2000) 'just pause and show the spash screen
        '    'must have already sent an email today.
        '    'MessageBox.Show("debug")
        '    End
        'End If

        ProcessIt()
        Application.Exit()

    End Sub


    Sub ProcessIt()

        Dim r As New Random
        Dim rn As Integer

        Dim sr As New StreamReader(TheAppDataPath & "Template.htm")
        Dim s As String = sr.ReadToEnd
        sr.Close()

        sr = New StreamReader(TheAppDataPath & "NextVacation.inf")
        'Dim nv As Date = sr.ReadLine
        Dim tmp As String = sr.ReadToEnd
        Dim nv As Date = GetNextDate(tmp)
        sr.Close()

        'get the following dates...
        Dim FollowingDates As String = GetFollowingDates(nv, tmp)

        Dim fType As String = ""

        Dim Suffix As String = ""
        If nv.Month = 11 Or nv.Month = 12 Then
            Suffix = "c" 'christmas in Nov and Dec
            fType = ".jpg"

        Else


            rn = r.Next(1, 10)

            If rn = 5 Then
                rn = r.Next(1, 10)
                If rn <= 5 Then
                    '10% of the time 'juj it up'
                    Suffix = "e"
                    fType = ".png"
                Else
                    '10% of the time 'juj it up'
                    Suffix = "d"
                    fType = ".png"
                End If
            Else
                'Default...
                Suffix = "a" 'Standard defualt
                fType = ".jpg"
            End If


        End If


        Dim n As Integer = DateDiff(DateInterval.Day, Now, nv)


        'debug... n = 9

        Dim sDays As String = ""

        If n >= 100 Then
            sDays = n.ToString
            GoTo SkipOut
        End If
        If n >= 10 Then
            sDays = "0" & n.ToString
            GoTo SkipOut
        End If
        If n < 10 Then
            sDays = "00" & n.ToString
        End If

SkipOut: '12.27.2022
        Dim Suffix1 As String = Suffix
        Dim Suffix2 As String = Suffix
        Dim Suffix3 As String = Suffix

        Dim fType1 As String = fType
        Dim fType2 As String = fType
        Dim fType3 As String = fType


        '*****************************************************************************************************************
        'Let's show zero the dog 1 out of every 10 times...
        '*****************************************************************************************************************
        If (Mid(sDays, 1, 1) = 0 Or Mid(sDays, 2, 1) = 0 Or Mid(sDays, 3, 1) = 0) And r.Next(1, 10) = 5 Then

            If Mid(sDays, 1, 1) = 0 Then
                Suffix1 = "Dog"
                fType1 = ".jpg"
            End If

            If Mid(sDays, 2, 1) = 0 Then
                Suffix2 = "Dog"
                fType2 = ".jpg"
            End If

            If Mid(sDays, 3, 1) = 0 Then
                Suffix3 = "Dog"
                fType3 = ".jpg"
            End If


        End If
        '*****************************************************************************************************************





        '****************************************************************
        '***  hide leading zero ***
        '****************************************************************
        If Mid(sDays, 1, 1) = "0" Then
            s = Replace(s, "<img src='http://www.schimsky.com/vacationcountdownemailer/NUM1' height='500' id='img1'>", "")
            's = Replace(s, "id='img1'", "id='img1' style='display:none'")
        End If

        If Mid(sDays, 2, 1) = "0" And Mid(sDays, 1, 1) = "0" Then
            s = Replace(s, "<img src='http://www.schimsky.com/vacationcountdownemailer/NUM2' height='500' id='img2'>", "")
            's = Replace(s, "id='img2'", "id='img2' style='display:none'")
        End If
        '****************************************************************
        '****************************************************************

        s = Replace(s, "NUM1", Mid(sDays, 1, 1) & Suffix1 & fType1)
        s = Replace(s, "NUM2", Mid(sDays, 2, 1) & Suffix2 & fType2)
        s = Replace(s, "NUM3", Mid(sDays, 3, 1) & Suffix3 & fType3)

        Dim iExtra As Integer = 0
        If sDays = "000" Then
            s = Replace(s, "yipee", "<img src='http://www.schimsky.com/vacationcountdownemailer/cc.gif' height='300'>")
            iExtra = 1
        End If

        If sDays = "007" Then
            s = Replace(s, "yipee", "<img src='http://www.schimsky.com/vacationcountdownemailer/oneweek.jpg' height='300'>")
        End If

        If iExtra = 0 Then
            s = Replace(s, "yipee", "")
        End If


        s = Replace(s, "FollowingDates", FollowingDates)


        'Print out to file, just for debugging....
        Dim sw As New StreamWriter(DailyPath, False)
            sw.Write(s)
            sw.Flush()
            sw.Close()

            'MessageBox.Show("debug..")
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
        'Smtp_Server.Credentials = New Net.NetworkCredential("sschimsky@tampabay.rr.com", "local") 'used with roadrunner cable
        'Smtp_Server.Credentials = New Net.NetworkCredential("steve@schimsky.com", "evabean2014")
        'Smtp_Server.Credentials = New Net.NetworkCredential("schimsky@outlook.com", "L0cal$0nly") 'used with outlook.com
        Smtp_Server.Credentials = New Net.NetworkCredential("schimsky@gmail.com", "kmltwiscpvrukyjh") 'used with GMail password generated by google security "app password" bypasses 2 factor auth

        'Smtp_Server.Host = "mail.twc.com" 'used with roadrunner cable
        'Smtp_Server.Host = "smtp.office365.com" 'outlook.com
        Smtp_Server.Host = "smtp.gmail.com" 'gmail
        'Smtp_Server.Host = "smtpout.secureserver.net" 'GoDaddy

        Smtp_Server.Port = 587 'used with outlook.com and  roadrunner cable
        Smtp_Server.EnableSsl = True 'used with outlook.com and roadrunner cable

        e_mail = New MailMessage()
        'e_mail.From = New MailAddress("sschimsky@tampabay.rr.com")
        'e_mail.From = New MailAddress("schimsky@outlook.com")
        e_mail.From = New MailAddress("schimsky@gmail.com")
        e_mail.To.Add("schimsky@gmail.com")
        e_mail.Subject = "VacationCountDownEmailer"
        e_mail.IsBodyHtml = True
        e_mail.Body = s

        Smtp_Server.Send(e_mail)


    End Sub




    Function GetNextDate(s) As Date

        'Get the next date from our .ini file

        Dim d As Date
        Dim td As Date
        Dim a = Split(s, vbCrLf)
        Dim tmp As String
        Dim i As Integer
        For i = 0 To UBound(a)
            tmp = a(i)
            If IsDate(tmp) Then
                td = tmp
                If td >= CDate(Now.ToShortDateString) Then
                    d = td
                    Return d
                End If
            End If
        Next

        If Not IsDate(d) Then
            d = DateAdd(DateInterval.Day, 900, Now)
        End If

        Return d

    End Function


    Function GetFollowingDates(d, s) As String

        Dim sReturn As String = ""

        Dim a = Split(s, vbCrLf)
        Dim tmp As String
        Dim i As Integer
        For i = 0 To UBound(a)
            tmp = a(i)

            Try
                If CDate(tmp) > d Then
                    Dim n As Integer = DateDiff(DateInterval.Day, Now, CDate(tmp))
                    sReturn = sReturn & tmp & " (" & n & " days) <br>"
                End If

            Catch ex As Exception

            End Try


        Next

        Return sReturn

    End Function


End Class
