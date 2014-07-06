Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.IO

Public Class Paypal

#Region "Send Money"
    Dim x As New CookieContainer
    Dim sessionstuff As New SessionData
    Dim sessionstuff2 As New SessionData
    Dim sessionstuff3 As New SessionData
    Public Structure SessionData
        Dim session As String
        Dim dispatch As String
        Dim context As String
    End Structure
    Function parselogin(ByVal source As String) As SessionData
        Dim data As SessionData
        Dim r1 As New Regex("SESSION=[a-zA-Z0-9-_]*")
        Dim matches1 As MatchCollection = Regex.Matches(source, r1.ToString, RegexOptions.Multiline)
        Dim sessionfix = matches1.Item(0).ToString.Replace("SESSION=", "")
        If matches1.Count > 0 Then data.session = sessionfix

        'SESSION=rXmm5weuhwCOdR8Xd9Un9kl03RgLIhZNXfi2Ef-;dispatch=453650344541fa3640b8dfba9cf6ea9f795b5f62d7c098aebd433f86039f4d2b">
        Dim r2 As New Regex("dispatch=[0-9A-Za-z-_]*")
        Dim matches2 As MatchCollection = Regex.Matches(source, r2.ToString, RegexOptions.Multiline)
        Dim dispatchfix = matches2.Item(0).ToString.Replace("dispatch=", "")
        If matches2.Count > 0 Then data.dispatch = dispatchfix

        Dim r3 As New Regex("CONTEXT=[a-zA-Z0-9-_;]*")
        Dim matches3 As MatchCollection = Regex.Matches(source, r3.ToString, RegexOptions.Multiline)
        Dim contextfix = matches3.Item(0).ToString.Replace("CONTEXT=", "")
        If matches3.Count > 0 Then data.context = contextfix
        Return data
    End Function
    Sub getlogin(ByVal ppusername As String, ByVal pppassword As String, ByVal dollaramount As String, ByVal centamount As String, ByVal recipient As String, ByVal note As String)
        Dim request As HttpWebRequest = HttpWebRequest.Create("https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_sm")
        With request
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "GET"

            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            sessionstuff = parselogin(dataresponse)

        End With

        Postlogin(ppusername, pppassword, dollaramount, centamount, recipient, note)
    End Sub
    Function Postlogin(ByVal ppusername As String, ByVal pppassword As String, ByVal dollaramount As String, ByVal centamount As String, ByVal recipient As String, ByVal note As String)
        Dim url = "https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_flow&SESSION=" & sessionstuff.session & "&dispatch=" & sessionstuff.dispatch
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        With request
            .Referer = "https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_sm"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "POST"
            Dim sb As New StringBuilder
            sb.Append("CONTEXT=" & sessionstuff.context)
            sb.Append("&login_email=" & ppusername)
            sb.Append("&login_password=" & pppassword)
            sb.Append("&login.x=Log In")
            sb.Append("&form_charset=UTF-8")

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(sb.ToString)
            .ContentLength = byteArray.Length
            Dim dataStream As Stream = .GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close() : dataStream.Dispose() : dataStream = Nothing

            Dim response As System.Net.HttpWebResponse = .GetResponse

            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            sessionstuff2 = parseafterlogin(dataresponse)
            If dataresponse.Contains("Cancel payment") Then
                '  Form1.ListBox1.Items.Add("Login Has Succeeded!" & Environment.NewLine)

            Else
                '  Form1.ListBox1.Items.Add("Login Has Failed!" & Environment.NewLine)
                Exit Function
            End If

        End With

        Postsendmoney1(ppusername, pppassword, dollaramount, centamount, recipient, note)

    End Function
    Function parseafterlogin(ByVal source As String) As SessionData
        Dim data As SessionData
        Dim r1 As New Regex("SESSION=[a-zA-Z0-9-_]*")
        Dim matches1 As MatchCollection = Regex.Matches(source, r1.ToString, RegexOptions.Multiline)
        Dim sessionfix = matches1.Item(0).ToString.Replace("SESSION=", "")
        If matches1.Count > 0 Then data.session = sessionfix

        Dim r3 As New Regex("CONTEXT=[a-zA-Z0-9-_;]*")
        Dim matches3 As MatchCollection = Regex.Matches(source, r3.ToString, RegexOptions.Multiline)
        Dim contextfix = matches3.Item(0).ToString.Replace("CONTEXT=", "")
        If matches3.Count > 0 Then data.context = contextfix

        Return data
    End Function
    Function Postsendmoney1(ByVal ppusername As String, ByVal pppassword As String, ByVal dollaramount As String, ByVal centamount As String, ByVal recipient As String, ByVal note As String)
        Dim url = "https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_flow&SESSION=" & sessionstuff2.session & "&dispatch=" & sessionstuff.dispatch
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        With request
            .Referer = "https://mobile.paypal.com/us/cgi-bin/webscr?cmd=_flow&SESSION=" & sessionstuff.session & "&dispatch=" & sessionstuff.dispatch
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "POST"
            Dim sb As New StringBuilder
            sb.Append("CONTEXT=" & sessionstuff2.context)
            sb.Append("&amount=" & dollaramount)
            sb.Append("&decimal_amount=" & centamount)
            sb.Append("&payment_recipient_alias=" & recipient)
            sb.Append("&note=" & note)
            sb.Append("&continue.x=Continue")
            sb.Append("&form_charset=UTF-8")

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(sb.ToString)
            .ContentLength = byteArray.Length
            Dim dataStream As Stream = .GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close() : dataStream.Dispose() : dataStream = Nothing

            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd

            sessionstuff3 = parseafterstartsending(dataresponse)
        End With

        FinalSendMoneyNow()

    End Function
    Function parseafterstartsending(ByVal source As String) As SessionData
        Dim data As SessionData
        Dim r1 As New Regex("SESSION=[a-zA-Z0-9-_]*")
        Dim matches1 As MatchCollection = Regex.Matches(source, r1.ToString, RegexOptions.Multiline)
        Dim sessionfix = matches1.Item(0).ToString.Replace("SESSION=", "")
        If matches1.Count > 0 Then data.session = sessionfix

        Dim r3 As New Regex("CONTEXT=[a-zA-Z0-9-_;]*")
        Dim matches3 As MatchCollection = Regex.Matches(source, r3.ToString, RegexOptions.Multiline)
        Dim contextfix = matches3.Item(0).ToString.Replace("CONTEXT=", "")
        If matches3.Count > 0 Then data.context = contextfix
        Return data
    End Function
    Function FinalSendMoneyNow()
        Dim url = "https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_flow&SESSION=" & sessionstuff3.session & "&dispatch=" & sessionstuff.dispatch
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        With request
            .Referer = "https://mobile.paypal.com/us/cgi-bin/webscr?cmd=_flow&SESSION=" & sessionstuff2.session & "&dispatch=" & sessionstuff.dispatch
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "POST"
            Dim sb As New StringBuilder
            sb.Append("CONTEXT=" & sessionstuff3.context)
            sb.Append("&send.x=Send Now")
            sb.Append("&form_charset=UTF-8")

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(sb.ToString)
            .ContentLength = byteArray.Length
            Dim dataStream As Stream = .GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close() : dataStream.Dispose() : dataStream = Nothing

            Dim response As System.Net.HttpWebResponse = .GetResponse

            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
        End With
        MsgBox("Money sent!")


    End Function
#End Region
#Region "Login To Normal Paypal"
#Region "Declarations"
    Public C As New CookieContainer
    Public Money As Double = 0
    Public Session As New SessionInfo
    Public PagesInSubscribers As Integer = 0
    Public PagesGoneThrough As Integer = 0
    Public NumberSubsGrabbed As Integer = 0
#End Region
#Region "Events"
    Public Event CustomerInfo(CustomerInformation As CustomerInformation)
    Public Event Stats(Status As String)
    Public Event SubsGrabbed(Subs As Integer)
    Public Event NewPageGrabbed(Source As String, TotalPages As String, CurrentPage As String)
    Public Event NewRecentPayment(RecentPayment As RecentPayment)
    Public Event CloseDB()
#End Region
#Region "Structures"
    Public Structure SessionInfo
        Dim DispatchVar
        Dim ContextVar
        Dim AuthVar
    End Structure
    Public Structure CustomerInformation
        Dim Name
        Dim Email
        Dim CustomerID
        Dim StartDate
        Dim Price
        Dim Status
    End Structure
    Public Structure RecentPayment
        Dim Name
        Dim Amount
        Dim Email
        Dim RecievedDate
        Dim Time
    End Structure
#End Region
#Region "Login To Paypal"
    Sub Login(ByVal ppusername As String, ByVal pppassword As String, DoWhatAfter As String)
        LoadLoginPage(ppusername, pppassword, DoWhatAfter)
    End Sub
    Sub LoadLoginPage(ByVal ppusername As String, ByVal pppassword As String, DoWhatAfter As String)
        RaiseEvent Stats("Thread {0} Trying to Login!" & vbNewLine)
        Dim request As HttpWebRequest = HttpWebRequest.Create("https://www.paypal.com/home")
        With request
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "GET"

            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
        End With
        PostLoginPage(ppusername, pppassword, DoWhatAfter)
    End Sub
    Sub PostLoginPage(ByVal ppusername As String, ByVal pppassword As String, DoWhatAfter As String)
        RaiseEvent Stats("Thread {0} Posting to login page." & vbNewLine)
        Dim url = "https://www.paypal.com/us/cgi-bin/webscr?cmd=_login-submit"
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        With request
            .Referer = "https://www.paypal.com/home"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "POST"
            Dim sb As New StringBuilder
            sb.Append("login_email=" & ppusername)
            sb.Append("&login_password=" & pppassword)
            sb.Append("&submit.x=Log In")
            sb.Append("&browser_name=Chrome")
            sb.Append("&browser_version=537.36")
            sb.Append("&browser_version_full=27.0.1453.94")
            sb.Append("&operating_system=Windows")
            sb.Append("&bp_mid=v%3D1%3Ba1%3Dna%7Ea2%3Dna%7Ea3%3Dna%7Ea4%3DMozilla%7Ea5%3DNetscape%7Ea6%3D5.0+%28Windows+NT+6.1%3B+WOW64%29+AppleWebKit%2F537.36+%28KHTML%2C+like+Gecko%29+Chrome%2F27.0.1453.94+Safari%2F537.36%7Ea7%3Dna%7Ea8%3Dna%7Ea9%3Dtrue%7Ea10%3Dna%7Ea11%3Dtrue%7Ea12%3DWin32%7Ea13%3Dna%7Ea14%3DMozilla%2F5.0+%28Windows+NT+6.1%3B+WOW64%29+AppleWebKit%2F537.36+%28KHTML%2C+like+Gecko%29+Chrome%2F27.0.1453.94+Safari%2F537.36%7Ea15%3Dtrue%7Ea16%3Dna%7Ea17%3DISO-8859-1%7Ea18%3Dwww.paypal.com%7Ea19%3Dna%7Ea20%3Dna%7Ea21%3Dna%7Ea22%3Dna%7Ea23%3D1024%7Ea24%3D768%7Ea25%3D32%7Ea26%3D768%7Ea27%3Dna%7Ea28%3Dna%7Ea29%3Dna%7Ea30%3Dna%7Ea31%3Dna%7Ea32%3Dna%7Ea33%3Dna%7Ea34%3Dna%7Ea35%3Dna%7Ea36%3Dna%7Ea37%3Dna%7Ea38%3Dna%7Ea39%3Dna%7Ea40%3Dna%7Ea41%3Dna%7Ea42%3Dna%7E")
            sb.Append("&bp_ks1=v%3D1%3Bl%3D9%3BDi0%3A75416Ui0%3A108Di1%3A238Ui1%3A108Di2%3A87Ui2%3A108Di3%3A95Ui3%3A108Di4%3A46Ui4%3A106Di5%3A60Ui5%3A69Di6%3A42Ui6%3A107Di7%3A75Ui7%3A104Di8%3A199Ui8%3A106Di9%3A86")
            sb.Append("&fso_enabled=11")
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(sb.ToString)
            .ContentLength = byteArray.Length
            Dim dataStream As Stream = .GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close() : dataStream.Dispose() : dataStream = Nothing

            Dim response As System.Net.HttpWebResponse = .GetResponse

            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            '  textbox.Text = dataresponse
            Dim r1 As New Regex("conds, <a href=""[A-Za-z0-9=_+-:?%&$#;@]*")
            Dim matches1 As MatchCollection = Regex.Matches(dataresponse, r1.ToString, RegexOptions.Multiline)
            Dim FinalURL = matches1.Item(0).ToString.Replace("conds, <a href=""", "")
            GetMainLoggedInPage(FinalURL, DoWhatAfter)

        End With
    End Sub
    Sub GetMainLoggedInPage(URL As String, DoWhatAfter As String)

        Dim request As HttpWebRequest = HttpWebRequest.Create(URL)
        With request
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "GET"

            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            RaiseEvent Stats("Thread {0} Logged In Successfully!" & vbNewLine)
            If DoWhatAfter = "FindSubscriberInfo" Then
                GetSubscribersPage()
            ElseIf DoWhatAfter = "CheckRecentTransactions" Then
                GrabHistoryPage()
            End If
        End With
        'conds, <a href="https://www.paypal.com/us/cgi-bin/webscr?cmd=%5flogin%2ddone&#x26;login&#x5f;access&#x3d;1371437866">click here</a> to re
        'https://www.paypal.com/us/cgi-bin/webscr?cmd=%5flogin%2ddone&#x26;login&#x5f;access&#x3d;1371437703
    End Sub
#End Region
#Region "Grab All Subscribers Ever"
    Function CollectSessionInfo(dataresponse As String)
        Dim r2 As New Regex("dispatch=[0-9A-Za-z-_]*")
        Dim matches2 As MatchCollection = Regex.Matches(dataresponse, r2.ToString, RegexOptions.Multiline)
        Dim dispatchfix = matches2.Item(0).ToString.Replace("dispatch=", "")
        Session.DispatchVar = dispatchfix

        Dim r3 As New Regex("CONTEXT""\svalue=""[a-zA-Z0-9-_;#&]*")
        Dim matches3 As MatchCollection = Regex.Matches(dataresponse, r3.ToString, RegexOptions.Multiline)
        Dim contextfix = matches3.Item(0).ToString.Replace("CONTEXT"" value=""", "")
        Session.ContextVar = contextfix

        Dim r4 As New Regex("auth""\stype=""hidden""\svalue=""[a-zA-Z0-9&%#;]*")
        Dim matches4 As MatchCollection = Regex.Matches(dataresponse, r4.ToString, RegexOptions.Multiline)
        Dim authfix = matches4.Item(0).ToString.Replace("auth"" type=""hidden"" value=""", "")
        Session.AuthVar = authfix
        RaiseEvent Stats("Thread {0} Grabbed Variables for Page " & PagesGoneThrough + 1 & "!" & vbNewLine)
    End Function
    Function FindAllEverSubscribers(source As String)
        Dim SubInAllCount As Integer = 0
        Dim r2 As New Regex("\([0-9]*\)")
        Dim matches2 As MatchCollection = Regex.Matches(source, r2.ToString, RegexOptions.Multiline)
        For Each Match As Match In matches2
            Dim active = Match.ToString.Replace("(", "")
            Dim active2 = active.Replace(")", "")
            If Not active2 = "" Then
                SubInAllCount += active2
            End If
        Next
        RaiseEvent Stats("Thread {0} Found Number of Subscribers to be " & SubInAllCount & vbNewLine)
        Return SubInAllCount
    End Function
    Function FindPagesNeeded(AllEverSubs As Integer)
        Dim pages As Integer = Math.Ceiling(AllEverSubs / 25)
        RaiseEvent Stats("Thread {0} Found Number of Pages to be " & pages & vbNewLine)
        Return pages
    End Function
    Function FindInformation2(dataresponse As String)

        Dim source = System.Net.WebUtility.HtmlDecode(dataresponse)

        'Dim CompleteRegex1 As New Regex(">(?<Name>[^<]*)<[^<]*<[^<]*<[^<]*<[^<]*Email:[^<]*<[^>]*>(?<email>[^<]*)<[^<]*<[^<]*<[^>]*>Customer\sID:[^<]*<[^>]*>(?<CustomerID>[^<]*)<[^<]*<[^<]*<[^>]*>Description:[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<DateStart>[a-zA-Z0-9&#;]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z]*)", RegexOptions.ExplicitCapture)
        Dim CompleteRegex2 As New Regex(">(?<Name>[^<]*)<[^<]*<[^<]*<[^<]*<[^<]*Email:[^<]*<[^>]*>(?<email>[^<]*)<[^<]*<[^<]*<[^>]*>Customer\sID:[^<]*<[^>]*>(?<CustomerID>[^<]*)<[^<]*<[^<]*<[^>]*>Description:[^<]*<[^<]*<[a-zA-Z0-9&#;]*/p><[^<]*<[^<]*<[^<]*<[^<]*<[^b]*balloonControl[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<DateStart>[^<]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z0-9]*)", RegexOptions.ExplicitCapture)
        '  Dim CompleteRegex As New Regex(">(?<Name>[^<]*)<[^<]*<[^<]*<[^<]*<[^<]*Email:[^<]*<[^>]*>(?<email>[^<]*)<[^<]*<[^<]*<[^>]*>Customer\sID:[^<]*<[^>]*>(?<CustomerID>[^<]*)<[^<]*<[^<]*<[^>]*>Description:([^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<DateStart>[a-zA-Z0-9&#;]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z]*)|[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<Date>[^<]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z0-9]*))", RegexOptions.ExplicitCapture)
        'Dim CompleteRegex As New Regex(">([^<]*)</a><div id=""[a-zA-Z0-9]*""\s[^E]*Email:\s*</strong>([^<]*)[^C]*Customer\sID:\s*</strong>([^<]*)[^D]*Description:\s*</strong>[^<]*</p></div></td><td\s*class=""desc""\sheaders=""description"">(<span\sclass=""nonjshide"">|)[^<]*(<a class=""balloonControl""\shref=""[0-9a-zA-Z#&;]*""\sid=""[a-zA-Z0-9#&;]*"">[^<]*</a></span><div\s*id=""[a-zA-Z0-9#&;]*""\sclass=""nonjsshow"">[a-zA-Z0-9#&;]*</div>|)</td><td\sclass=""date""\sheaders=""start"">([^<]*)</td><td\s*class=""date""\s*headers=""end""><.td><td\s*headers=""lastbill"">([^<]*)</td><td\sheaders=""status"">([a-zA-Z0-9]*)")
        ' Dim CompleteRegex As New Regex(">([^<]*)</a><div\sid=""[a-zA-Z0-9]*""\s[^E]*Email:\s*</strong>([^<]*)[^C]*Customer\sID:\s*</strong>([^<]*)[^D]*Description:\s*</strong>[^<]*</p></div></td><td\s*class=""desc""\sheaders=""description"">[^<]*</td><td\sclass=""date""\sheaders=""start"">([^<]*)</td><td\s*class=""date""\s*headers=""end""><.td><td\s*headers=""lastbill"">([^<]*)</td><td\sheaders=""status"">([a-zA-Z0-9]*)")
        '   Dim CompleteRegex As New Regex(">([a-zA-Z0-9&#;…]*)</a><div id=""name[0-9]*""\sclass=""nonjsshow""><p class=""bcontent""><strong>Email: </strong>([a-zA-Z0-9-_&#;…]*)</p><p class=""bcontent""><strong>Customer ID: </strong>([a-zA-Z0-9-_#$%&;]*)</p><p class=""bcontent""><strong>Description:  </strong>[a-zA-Z0-9&#;]*</p></div></td><td class=""desc"" headers=""description"">[a-zA-Z0-9&#;]*</td><td class=""date"" headers=""start"">[a-zA-Z0-9&#;]*</td><td class=""date"" headers=""end""></td><td headers=""lastbill"">([a-zA-Z0-9-_&#;]*)</td><td headers=""status"">([a-zA-Z]*)")
        Dim CompleteMatches1 As MatchCollection = Regex.Matches(source, CompleteRegex2.ToString, RegexOptions.ExplicitCapture)
        For Each Match As Match In CompleteMatches1
            Dim Customer As New CustomerInformation
            Customer.Name = System.Net.WebUtility.HtmlDecode(Match.Groups(1).Value)
            Customer.Email = System.Net.WebUtility.HtmlDecode(Match.Groups(2).Value)
            Customer.CustomerID = System.Net.WebUtility.HtmlDecode(Match.Groups(3).Value)
            Customer.StartDate = System.Net.WebUtility.HtmlDecode(Match.Groups(4).Value)
            Customer.Price = System.Net.WebUtility.HtmlDecode(Match.Groups(5).Value)
            Customer.Status = System.Net.WebUtility.HtmlDecode(Match.Groups(6).Value)
            RaiseEvent CustomerInfo(Customer)
            NumberSubsGrabbed += 1
            RaiseEvent SubsGrabbed(NumberSubsGrabbed)
        Next
        FindInformation2(dataresponse)


        'Dim r1 As New Regex("Email:\s</strong>[a-zA-Z0-9-_%$&#;:@\.]*")
        'Dim matches1 As MatchCollection = Regex.Matches(source, r1.ToString, RegexOptions.Multiline)
        'For Each Match In matches1
        '    Dim Email = System.Net.WebUtility.HtmlDecode(Match.ToString.Replace("Email: </strong>", ""))

        '    TextBox.AppendText(Email & Environment.NewLine)

        'Next
    End Function
    Function FindInformation(dataresponse As String)

        Dim source = System.Net.WebUtility.HtmlDecode(dataresponse)

        Dim CompleteRegex As New Regex(">(?<Name>[^<]*)<[^<]*<[^<]*<[^<]*<[^<]*Email:[^<]*<[^>]*>(?<email>[^<]*)<[^<]*<[^<]*<[^>]*>Customer\sID:[^<]*<[^>]*>(?<CustomerID>[^<]*)<[^<]*<[^<]*<[^>]*>Description:[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<DateStart>[a-zA-Z0-9&#;]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z]*)", RegexOptions.ExplicitCapture)
        ' Dim CompleteRegex2 As New Regex(">(?<Name>[^<]*)<[^<]*<[^<]*<[^<]*<[^<]*Email:[^<]*<[^>]*>(?<email>[^<]*)<[^<]*<[^<]*<[^>]*>Customer\sID:[^<]*<[^>]*>(?<CustomerID>[^<]*)<[^<]*<[^<]*<[^>]*>Description:[^<]*<[^<]*<[a-zA-Z0-9&#;]*/p><[^<]*<[^<]*<[^<]*<[^<]*<[^b]*balloonControl[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<DateStart>[^<]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z0-9]*)", RegexOptions.ExplicitCapture)
        '  Dim CompleteRegex As New Regex(">(?<Name>[^<]*)<[^<]*<[^<]*<[^<]*<[^<]*Email:[^<]*<[^>]*>(?<email>[^<]*)<[^<]*<[^<]*<[^>]*>Customer\sID:[^<]*<[^>]*>(?<CustomerID>[^<]*)<[^<]*<[^<]*<[^>]*>Description:([^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<DateStart>[a-zA-Z0-9&#;]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z]*)|[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<Date>[^<]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z0-9]*))", RegexOptions.ExplicitCapture)
        'Dim CompleteRegex As New Regex(">([^<]*)</a><div id=""[a-zA-Z0-9]*""\s[^E]*Email:\s*</strong>([^<]*)[^C]*Customer\sID:\s*</strong>([^<]*)[^D]*Description:\s*</strong>[^<]*</p></div></td><td\s*class=""desc""\sheaders=""description"">(<span\sclass=""nonjshide"">|)[^<]*(<a class=""balloonControl""\shref=""[0-9a-zA-Z#&;]*""\sid=""[a-zA-Z0-9#&;]*"">[^<]*</a></span><div\s*id=""[a-zA-Z0-9#&;]*""\sclass=""nonjsshow"">[a-zA-Z0-9#&;]*</div>|)</td><td\sclass=""date""\sheaders=""start"">([^<]*)</td><td\s*class=""date""\s*headers=""end""><.td><td\s*headers=""lastbill"">([^<]*)</td><td\sheaders=""status"">([a-zA-Z0-9]*)")
        ' Dim CompleteRegex As New Regex(">([^<]*)</a><div\sid=""[a-zA-Z0-9]*""\s[^E]*Email:\s*</strong>([^<]*)[^C]*Customer\sID:\s*</strong>([^<]*)[^D]*Description:\s*</strong>[^<]*</p></div></td><td\s*class=""desc""\sheaders=""description"">[^<]*</td><td\sclass=""date""\sheaders=""start"">([^<]*)</td><td\s*class=""date""\s*headers=""end""><.td><td\s*headers=""lastbill"">([^<]*)</td><td\sheaders=""status"">([a-zA-Z0-9]*)")
        '   Dim CompleteRegex As New Regex(">([a-zA-Z0-9&#;…]*)</a><div id=""name[0-9]*""\sclass=""nonjsshow""><p class=""bcontent""><strong>Email: </strong>([a-zA-Z0-9-_&#;…]*)</p><p class=""bcontent""><strong>Customer ID: </strong>([a-zA-Z0-9-_#$%&;]*)</p><p class=""bcontent""><strong>Description:  </strong>[a-zA-Z0-9&#;]*</p></div></td><td class=""desc"" headers=""description"">[a-zA-Z0-9&#;]*</td><td class=""date"" headers=""start"">[a-zA-Z0-9&#;]*</td><td class=""date"" headers=""end""></td><td headers=""lastbill"">([a-zA-Z0-9-_&#;]*)</td><td headers=""status"">([a-zA-Z]*)")
        Dim CompleteMatches As MatchCollection = Regex.Matches(source, CompleteRegex.ToString, RegexOptions.ExplicitCapture)
        For Each Match As Match In CompleteMatches
            Dim Customer As New CustomerInformation
            Customer.Name = System.Net.WebUtility.HtmlDecode(Match.Groups(1).Value)
            Customer.Email = System.Net.WebUtility.HtmlDecode(Match.Groups(2).Value)
            Customer.CustomerID = System.Net.WebUtility.HtmlDecode(Match.Groups(3).Value)
            Customer.StartDate = System.Net.WebUtility.HtmlDecode(Match.Groups(4).Value)
            Customer.Price = System.Net.WebUtility.HtmlDecode(Match.Groups(5).Value)
            Customer.Status = System.Net.WebUtility.HtmlDecode(Match.Groups(6).Value)
            '  RaiseEvent Stats("Information Grabbed for " & Customer.Email & vbNewLine)
            RaiseEvent CustomerInfo(Customer)
            NumberSubsGrabbed += 1
            RaiseEvent SubsGrabbed(NumberSubsGrabbed)
            'textbox.AppendText(Customer.Status & ", " & Customer.Name & ", " & Customer.Email & ", " & Customer.CustomerID & ", " & Customer.StartDate & ", " & Customer.Price & vbNewLine)
        Next


        'Dim r1 As New Regex("Email:\s</strong>[a-zA-Z0-9-_%$&#;:@\.]*")
        'Dim matches1 As MatchCollection = Regex.Matches(source, r1.ToString, RegexOptions.Multiline)
        'For Each Match In matches1
        '    Dim Email = System.Net.WebUtility.HtmlDecode(Match.ToString.Replace("Email: </strong>", ""))

        '    TextBox.AppendText(Email & Environment.NewLine)

        'Next
    End Function
    Sub GetSubscribersPage()
        Dim request As HttpWebRequest = HttpWebRequest.Create("https://www.paypal.com/cgi-bin/customerprofileweb?cmd=%5fmerchant%2dhub")
        With request
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "GET"

            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            RaiseEvent Stats("Thread {0} Grabbed Page 1 of Subscribers!" & vbNewLine)
            RaiseEvent NewPageGrabbed(dataresponse, PagesInSubscribers, PagesGoneThrough)
            ' FindInformation(dataresponse)
            Dim allsubs As Integer = FindAllEverSubscribers(dataresponse)
            PagesInSubscribers = FindPagesNeeded(allsubs)

            CollectSessionInfo(dataresponse)

            '  GetAllActiveSubscribers(dispatchfix, contextfix, authfix, textbox)
        End With
        PagesGoneThrough += 1
        GetAllSubscribersNextPage()
    End Sub
    Sub GetAllSubscribersNextPage()
        Dim request As HttpWebRequest = HttpWebRequest.Create("https://www.paypal.com/us/cgi-bin/webscr?cmd=_flow&dispatch=" & Session.DispatchVar)
        With request
            .Referer = "http://www.google.com"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "POST"
            Dim sb As New StringBuilder
            sb.Append("CONTEXT=" & System.Net.WebUtility.HtmlDecode(Session.ContextVar))
            sb.Append("&light_box_submit=none")
            sb.Append("&count=25")
            sb.Append("&filter_level_two=filter_all")
            sb.Append("&filter_level_three=filter_five_days")
            sb.Append("&next=Next page")
            sb.Append("&auth=" & System.Net.WebUtility.HtmlDecode(Session.AuthVar))
            sb.Append("&form_charset=UTF-8")
            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(sb.ToString)
            .ContentLength = byteArray.Length
            Dim dataStream As Stream = .GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close() : dataStream.Dispose() : dataStream = Nothing

            Dim response As System.Net.HttpWebResponse = .GetResponse

            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            RaiseEvent Stats("Thread {0} Grabbed Page " & PagesGoneThrough & " of Subscribers!" & vbNewLine)
            RaiseEvent NewPageGrabbed(dataresponse, PagesInSubscribers, PagesGoneThrough)
            ' FindInformation(dataresponse)



            CollectSessionInfo(dataresponse)
        End With

        PagesGoneThrough += 1
        If PagesGoneThrough >= PagesInSubscribers Then    ''''''''''''''''' ORIGINALLY +1
            RaiseEvent CloseDB()
        Else
            GetAllSubscribersNextPage()
        End If
    End Sub
#End Region
#Region "Grab Recent Transactions"
    Sub GrabHistoryPage()
        Dim request As HttpWebRequest = HttpWebRequest.Create("https://history.paypal.com/us/cgi-bin/webscr?cmd=_history")
        With request
            .Referer = "https://www.paypal.com/webapps/hub/"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "GET"

            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            RaiseEvent Stats("Thread {0} Grabbed Recent History Page!" & vbNewLine)
            Dim r As New Regex("dateInfo"">[^>]*>\s*([a-zA-Z0-9#&;]*)[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>([a-zA-Z0-9#&; ]*)[^>]*>[^>]*>[^>]*>[^>]*>\s*([a-zA-Z0-9&;#]*)[^>]*>[^>]*[^>]*>[^>]*>[^>]*>[^>]*>([a-zA-Z]*)[^>]*>[^>]*>[^>]*><a href=""([^""]*)")
            Dim m As MatchCollection = Regex.Matches(dataresponse, r.ToString, RegexOptions.Multiline)
            For Each Match As Match In m
                If System.Net.WebUtility.HtmlDecode(Match.Groups(4).Value) = "Completed" Then
                    GrabDetails(System.Net.WebUtility.HtmlDecode(Match.Groups(3).Value), Match.Groups(5).Value)
                End If
                If m.Count = Match.Index - 1 Then
                    RaiseEvent CloseDB()
                End If
            Next

        End With
    End Sub
    Sub GrabDetails(name As String, url As String)
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        With request
            .Referer = "https://history.paypal.com/us/cgi-bin/webscr?cmd=_history"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "GET"

            Dim response As System.Net.HttpWebResponse = .GetResponse
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
            Dim dataresponse As String = sr.ReadToEnd
            RaiseEvent Stats("Thread {0} Grabbed Buyer History For " & name & vbNewLine)
            Dim r As New Regex("Buyer Name:[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>([a-zA-Z0-9&#;]*)[^>]*>[^>]*>[^>]*>[^>]*>Buyer Email:[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>([a-zA-Z0-9#&;]*)[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>Total amount:[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>([0-9a-zA-Z#&;]*)")
            Dim m As MatchCollection = Regex.Matches(dataresponse, r.ToString, RegexOptions.Multiline)
            If m.Count = 0 Then
            Else
                Dim Payment As New RecentPayment
                Payment.Name = System.Net.WebUtility.HtmlDecode(m(0).Groups(1).Value)
                Payment.Email = System.Net.WebUtility.HtmlDecode(m(0).Groups(2).Value)
                Payment.Amount = System.Net.WebUtility.HtmlDecode(m(0).Groups(3).Value)

                Dim r1 As New Regex("Date:[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>([a-zA-Z0-9&#;]*)[^>]*>[^>]*>[^>]*>[^>]*>Time:[^>]*>[^>]*>[^>]*>[^>]*>[^>]*>([a-zA-Z0-9#&;]*)")
                Dim m1 As MatchCollection = Regex.Matches(dataresponse, r1.ToString, RegexOptions.Multiline)

                Payment.RecievedDate = System.Net.WebUtility.HtmlDecode(m1(0).Groups(1).Value)
                Payment.Time = System.Net.WebUtility.HtmlDecode(m1(0).Groups(2).Value)
                RaiseEvent NewRecentPayment(Payment)
            End If
        End With
    End Sub

#End Region
#End Region
#Region "Check Email For Paypal"
#Region "Events"
    Public Event SendProxiesTo(email As String, proxies As Integer)
#End Region
    Sub CheckForPaypalEmails(Host As String, LoginName As String, Password As String, Port As String, type As String)
        Try
            If Not My.Computer.FileSystem.FileExists("seenEmailsFromPaypal.txt") Then
                Dim fs As FileStream = File.Create("seenEmailsFromPaypal.txt")
                fs.Close()
            End If
            Dim mailman As New Chilkat.MailMan()

            '  Any string argument automatically begins the 30-day trial.
            Dim success As Boolean
            success = mailman.UnlockComponent("MAIL12345678_762D6411G83F")
            mailman.MailHost = Host

            '  Set the POP3 login/password.
            mailman.PopUsername = LoginName
            mailman.PopPassword = Password
            '  Copy the all email from the user's POP3 mailbox
            '  into a bundle object.  The email remains on the server.
            Dim saSeenUidls As New Chilkat.StringArray()
            success = saSeenUidls.LoadFromFile("seenEmailsFromPaypal.txt")
            If (success <> True) Then
                RaiseEvent Stats("failed to load seenEmailsFromPaypal.txt" & vbNewLine)
                Exit Sub
            End If


            '  Get the complete list of UIDLs on the mail server.
            Dim saUidls As Chilkat.StringArray
            saUidls = mailman.GetUidls()

            If (saUidls Is Nothing) Then
                MsgBox(mailman.LastErrorText)
                Exit Sub
            End If


            '  Create a new string array object (it's an object, not an actual array)
            '  and add the UIDLs from saUidls that aren't already seen.
            Dim saUnseenUidls As New Chilkat.StringArray()

            Dim i As Long
            Dim n As Long
            n = saUidls.Count
            For i = 0 To n - 1
                If (saSeenUidls.Contains(saUidls.GetString(i)) <> True) Then
                    saUnseenUidls.Append(saUidls.GetString(i))
                End If

            Next


            If (saUnseenUidls.Count = 0) Then
                RaiseEvent Stats("No Unread Emails from Paypal at " & TimeOfDay & vbNewLine)
                mailman.Pop3EndSession()
            Exit Sub
            End If


            '  Download in full the unseen emails:
            Dim bundle As Chilkat.EmailBundle

            bundle = mailman.FetchMultiple(saUnseenUidls)
            If (bundle Is Nothing) Then
                MsgBox(mailman.LastErrorText)

                Exit Sub
            End If


            Dim email As Chilkat.Email
            For i = 0 To bundle.MessageCount - 1


                email = bundle.GetEmail(i)

                If email.Subject.Contains("You received a payment") Then
                    getinfo(email.Body)
                    mailman.Pop3EndSession()
                Else
                    saUidls.SaveToFile("seenEmailsFromPaypal.txt")
                    mailman.Pop3EndSession()
                End If
            Next
            saUidls.SaveToFile("seenEmailsFromPaypal.txt")
            mailman.Pop3EndSession()
        Catch ex As Exception
            RaiseEvent Stats(ex.Message & Environment.NewLine)
        End Try
    End Sub
    Sub getinfo(body As String)
        Try
            Dim good As String = ""
            Dim good1 As String = ""
            Dim good2 As String = ""
            Dim good3 As String = ""
            Dim newgood As String = ""
            Dim r As String = ""
            Dim custname As New System.Text.RegularExpressions.Regex("Customer name:[A-Za-z-_]*")
            Dim custnamematches As MatchCollection = custname.Matches(body)
            For Each match As Match In custnamematches
                good = match.ToString.Replace("Customer name:", "")
            Next
            Dim custemail As New System.Text.RegularExpressions.Regex("Customer email:[A-Za-z-_@\.]*")
            Dim custemailmatches As MatchCollection = custemail.Matches(body)
            For Each match As Match In custemailmatches
                good1 = match.ToString.Replace("Customer email:", "")
            Next
            Dim custpaymentID As New System.Text.RegularExpressions.Regex("Profile ID:[0-9A-Za-z-_@\.]*")
            Dim custpaymentIDmatches As MatchCollection = custpaymentID.Matches(body)
            For Each match As Match In custpaymentIDmatches
                good2 = match.ToString.Replace("Profile ID:", "")
            Next
            Dim custpaideachtime As New System.Text.RegularExpressions.Regex("Amount paid each time:[0-9A-Za-z-_@\.$ ]*")
            Dim custpaideachtimematches As MatchCollection = custpaideachtime.Matches(body)
            For Each match As Match In custpaideachtimematches
                good3 = match.ToString.Replace("Amount paid each time:", "")
                newgood = good3.Replace("$", "")
                If good3.Contains("27") Then
                    r = newgood.Replace("27.00", "30")
                ElseIf good3.Contains("9") Then
                    r = newgood.Replace("9.00", "10")
                ElseIf good3.Contains("45") Then
                    r = newgood.Replace("45.00", "50")
                ElseIf good3.Contains("85") Then
                    r = newgood.Replace("85.00", "100")
                End If

            Next
            RaiseEvent SendProxiesTo(good1, r)
            ' MakeInfoFile(good, good1, good2, good3) $10.00 USD
        Catch ex As Exception
            RaiseEvent Stats(ex.Message & Environment.NewLine)
        End Try
    End Sub
    Sub MakeInfoFile(Name As String, Email As String, ProfileID As String, AmountPaid As String)
        Try
            Dim CurrentDir = My.Application.Info.DirectoryPath
            If (Not System.IO.Directory.Exists(CurrentDir & "\Profiles")) Then
                System.IO.Directory.CreateDirectory(CurrentDir & "\Profiles")
            End If
            If (Not System.IO.Directory.Exists(CurrentDir & "\Profiles\" & Name)) Then
                System.IO.Directory.CreateDirectory(CurrentDir & "\Profiles\" & Name)
                Dim UserDir As String = CurrentDir & "\Profiles\" & Name
                Dim theFile As FileStream = File.Create(UserDir & "\Proxies.txt")
                Dim theFile1 As FileStream = File.Create(UserDir & "\Information.txt")
                Dim theFile2 As FileStream = File.Create(UserDir & "\HasProxies.txt")
                theFile.Close()
                theFile1.Close()
                theFile2.Close()
                Dim objWriter As New System.IO.StreamWriter(UserDir & "\HasProxies.txt", False)
                objWriter.WriteLine("False")
                objWriter.Close()
                Dim objWriter1 As New System.IO.StreamWriter(UserDir & "\Information.txt", False)
                RaiseEvent Stats("Creating Profile for " & Name & " Who Paid " & AmountPaid & Environment.NewLine)
                objWriter1.WriteLine("[NAME]" & Name & "[NAME]")
                objWriter1.WriteLine("[EMAIL]" & Email & "[EMAIL]")
                objWriter1.WriteLine("[PROFILEID]" & ProfileID & "[PROFILEID]")
                objWriter1.WriteLine("[AMOUNTPAID]" & AmountPaid & "[AMOUNTPAID]")
                If AmountPaid.Contains("10.00") Then
                    objWriter1.WriteLine("[NUMBEROFPROXIES]10[NUMBEROFPROXIES]")
                ElseIf AmountPaid.Contains("20.00") Then
                    objWriter1.WriteLine("[NUMBEROFPROXIES]20[NUMBEROFPROXIES]")
                ElseIf AmountPaid.Contains("27.00") Then
                    objWriter1.WriteLine("[NUMBEROFPROXIES]30[NUMBEROFPROXIES]")
                ElseIf AmountPaid.Contains("46.00") Then
                    objWriter1.WriteLine("[NUMBEROFPROXIES]50[NUMBEROFPROXIES]")
                ElseIf AmountPaid.Contains("85.00") Then
                    objWriter1.WriteLine("[NUMBEROFPROXIES]100[NUMBEROFPROXIES]")
                End If
                objWriter1.WriteLine("[TIMES]+1[TIMES]")
                objWriter1.Close()
            Else
                Dim UserDir As String = CurrentDir & "\Profiles\" & Name
                Dim objWriter1 As New System.IO.StreamWriter(UserDir & "\Information.txt", True)
                objWriter1.WriteLine("[TIMES]+1[TIMES]")
                objWriter1.Close()
            End If
            RaiseEvent Stats("Finished Creating Profile for " & Name & Environment.NewLine)
        Catch ex As Exception
            RaiseEvent Stats(ex.Message & Environment.NewLine)
        End Try
    End Sub
    Sub SendProxiesToCustomer(email As String, proxies As Integer)
        '  The mailman object is used for sending and receiving email.
        Dim mailman As New Chilkat.MailMan()

        '  Any string argument automatically begins the 30-day trial.
        Dim success As Boolean
        success = mailman.UnlockComponent("MAIL12345678_762D6411G83F")
        '  Set the SMTP server.
        mailman.SmtpHost = "smtp.chilkatsoft.com"

        '  Set the SMTP login/password (if required)
        mailman.SmtpUsername = "myUsername"
        mailman.SmtpPassword = "myPassword"

        '  Create a new email object
        Dim email As New Chilkat.Email()

        email.Subject = "This is a test"
        email.Body = "This is a test"
        email.From = "Chilkat Support <support@chilkatsoft.com>"
        email.AddTo("Chilkat Admin", "admin@chilkatsoft.com")
        '  To add more recipients, call AddTo, AddCC, or AddBcc once per recipient.

        '  Call SendEmail to connect to the SMTP server and send.
        '  The connection (i.e. session) to the SMTP server remains
        '  open so that subsequent SendEmail calls may use the
        '  same connection.
        success = mailman.SendEmail(email)
        If (success <> True) Then
            TextBox1.Text = TextBox1.Text & mailman.LastErrorText & vbCrLf
            Exit Sub
        End If


        '  Some SMTP servers do not actually send the email until
        '  the connection is closed.  In these cases, it is necessary to
        '  call CloseSmtpConnection for the mail to be  sent.
        '  Most SMTP servers send the email immediately, and it is
        '  not required to close the connection.  We'll close it here
        '  for the example:
        success = mailman.CloseSmtpConnection()
        If (success <> True) Then
            MsgBox("Connection to SMTP server not closed cleanly.")
        End If


        MsgBox("Mail Sent!")
    End Sub
#End Region
End Class
