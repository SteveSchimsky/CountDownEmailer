Imports System.IO
Imports System.ComponentModel
Imports System.Net.Mail
Imports System.Drawing

Public Class Form1

    Dim Seconds As Integer = 0
    Dim DailyPath As String = Application.StartupPath & "\Data\Daily.htm"
    Dim TheAppDataPath As String = Application.StartupPath & "\Data\"
    Dim ImagePath As String = Application.StartupPath & "\Data\Images\"


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Show()
        Application.DoEvents()
        'MsgBox("debug")
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


        Dim myAltView As AlternateView


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
        s = Replace(s, "FollowingDates", FollowingDates)

        Dim fType As String = ""

        Dim Suffix As String = ""
        If nv.Month = 11 Or nv.Month = 12 Then
            Suffix = "c" 'christmas in Nov and Dec
            fType = ".jpg"

        Else


            rn = r.Next(1, 10)

            If rn = 5 Or rn = 6 Or rn = 7 Then
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
        'n = 999

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

        Dim DoFirst As Boolean = True
        Dim DoSecond As Boolean = True

        If Mid(sDays, 1, 1) = "0" Then
            'Hide it...
            s = Replace(s, "<img src='cid:111' height='500' id='img1'>", "")
            DoFirst = False
        End If

        If Mid(sDays, 2, 1) = "0" And Mid(sDays, 1, 1) = "0" Then
            'Hide it...
            s = Replace(s, "<img src='cid:222' height='500' id='img2'>", "")
            DoSecond = False
        End If
        '****************************************************************
        '****************************************************************



        Dim iExtra As Integer = 0
        If sDays = "000" Then
            'create an image tag that will be referenced by the embedded image
            s = Replace(s, "yipee", "<img src='cid:Yipee' height='500' id='img2'>")
            iExtra = 1
        End If

        If sDays = "007" Then
            'create an image tag that will be referenced by the embedded image
            s = Replace(s, "yipee", "<img src='cid:OneWeek' height='300' id='img2'>")
        End If


        If iExtra = 0 Then
            'Hide the 'yipee'
            s = Replace(s, "yipee", "")
        End If




        '*********************************************************************************************************
        'Only need this once...
        myAltView = AlternateView.CreateAlternateViewFromString(s, New System.Net.Mime.ContentType("text/html"))
        '*********************************************************************************************************


        Dim LR As LinkedResource = CreateLinkedResource("HEADER111", ImagePath & "HEADER.JPG", System.Net.Mime.MediaTypeNames.Image.Jpeg)
        myAltView.LinkedResources.Add(LR)


        If sDays = "000" Then
            'Add the Yipee We're there! image....
            LR = CreateLinkedResource("Yipee", ImagePath & "cc.gif", System.Net.Mime.MediaTypeNames.Image.Gif)
            myAltView.LinkedResources.Add(LR)
        End If


        If sDays = "007" Then
            'Add the 1 Week to Go! image...
            LR = CreateLinkedResource("OneWeek", ImagePath & "oneweek.jpg", System.Net.Mime.MediaTypeNames.Image.Jpeg)
            myAltView.LinkedResources.Add(LR)
        End If


        If DoFirst = True Then
            Dim Name1 As String = Mid(sDays, 1, 1) & Suffix1 & fType1
            LR = CreateLinkedResource("111", ImagePath & Name1, System.Net.Mime.MediaTypeNames.Image.Jpeg)
            myAltView.LinkedResources.Add(LR)
        End If


        If DoSecond = True Then
            Dim Name2 As String = Mid(sDays, 2, 1) & Suffix2 & fType2
            LR = CreateLinkedResource("222", ImagePath & Name2, System.Net.Mime.MediaTypeNames.Image.Jpeg)
            myAltView.LinkedResources.Add(LR)
        End If


        Dim Name3 As String = Mid(sDays, 3, 1) & Suffix3 & fType3
        LR = CreateLinkedResource("333", ImagePath & Name3, System.Net.Mime.MediaTypeNames.Image.Jpeg)
        myAltView.LinkedResources.Add(LR)



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
        e_mail.AlternateViews.Add(myAltView)

        'e_mail.From = New MailAddress("sschimsky@tampabay.rr.com")
        'e_mail.From = New MailAddress("schimsky@outlook.com")
        e_mail.From = New MailAddress("schimsky@gmail.com")
        e_mail.To.Add("schimsky@gmail.com")
        e_mail.Subject = "VacationCountDownEmailer"
        e_mail.IsBodyHtml = True
        e_mail.Body = s


        Smtp_Server.Send(e_mail)
        'MsgBox("done")

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



    Function CreateLinkedResource(PicID As String, PathToPIC As String, MediaTypeImage As String) As LinkedResource



        'GRAB IMAGE FROM FILE AND PUT IN MEMORY STREAM
        Dim myImage As Image
        myImage = Image.FromFile(PathToPIC) 'arraylist of my images

        Dim IC As ImageConverter
        Dim myImageData() As Byte

        IC = New ImageConverter
        myImageData = DirectCast(IC.ConvertTo(myImage, GetType(Byte())), Byte())
        Dim myStream As New MemoryStream(myImageData)

        'CREATE ALT VIEW
        Dim myLinkedResource As LinkedResource
        'CREATE LINKED RESOURCE FOR ALT VIEW
        myLinkedResource = New LinkedResource(myStream, MediaTypeImage)
        ''SET CONTENTID SO HTML CAN REFERENCE CORRECTLY
        myLinkedResource.ContentId = PicID 'this must match in the HTML of the message body
        '******************************************************************************
        'added by steve on 9.7.2022 ..
        myLinkedResource.ContentType.MediaType = MediaTypeImage 'System.Net.Mime.MediaTypeNames.Image.Jpeg
        myLinkedResource.TransferEncoding = System.Net.Mime.TransferEncoding.Base64
        myLinkedResource.ContentType.Name = PicID & "-Name"
        '******************************************************************************
        ''ADD LINKED RESOURCE TO ALT VIEW, AND ADD ALT VIEW TO MESSAGE

        Return myLinkedResource


    End Function




End Class
