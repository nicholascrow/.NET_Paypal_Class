Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.IO

Public Class Paypal
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
    Public Structure SessionData 'Structure containing the current Session Data
        Dim session As String
        Dim dispatch As String
        Dim context As String
    End Structure
#End Region

#Region "Send Money"
#Region "SendMoneySessions"
    Dim LoginSessionData As New SessionData ''''//Before Login//
    Dim PostLoginSessionData As New SessionData ''''//After Login//  Paypal uses sessions to see whether users are logged in or not.
    Dim PostSendMoneySessionData As New SessionData '''' //After Sending Money//
    Private Function ParseLoginData(ByVal source As String) As SessionData 'Use Regular Expressions to parse for Session, Dispatch, and Context, before logging in.
        Dim data As SessionData
        Dim r1 As New Regex("SESSION=[a-zA-Z0-9-_]*")
        Dim matches1 As MatchCollection = Regex.Matches(source, r1.ToString, RegexOptions.Multiline)
        Dim sessionfix = matches1.Item(0).ToString.Replace("SESSION=", "")
        If matches1.Count > 0 Then data.session = sessionfix


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
    Protected Function ParseSessionAfterLogin(ByVal source As String) As SessionData 'Grabs Session information again after logging in, because it changes.
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
    Private Function ParseFinalSendMoneyPage(ByVal source As String) As SessionData 'This is the session data after the money to be sent has already been implemented into paypal. The only thing left to do is press the send button.
        Dim data As SessionData = Nothing
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
#End Region

    Public Sub LoginAndSendMoney(ByVal ppusername As String, ByVal pppassword As String, ByVal dollaramount As String, ByVal centamount As String, ByVal recipient As String, ByVal note As String) 'This function starts the process of sending money.
        'This section loads the main paypal page, grabbing the cookie needed to login.
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
            LoginSessionData = ParseLoginData(dataresponse)

        End With

        PostToLoginPage(ppusername, pppassword, dollaramount, centamount, recipient, note)
    End Sub
    Protected Sub PostToLoginPage(ByVal ppusername As String, ByVal pppassword As String, ByVal dollaramount As String, ByVal centamount As String, ByVal recipient As String, ByVal note As String)
        'This function posts the user's login data directly to the paypal login page.
        Dim url = "https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_flow&SESSION=" & LoginSessionData.session & "&dispatch=" & LoginSessionData.dispatch
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        With request
            .Referer = "https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_sm"
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "POST"
            Dim sb As New StringBuilder
            sb.Append("CONTEXT=" & LoginSessionData.context)
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
            PostLoginSessionData = ParseSessionAfterLogin(dataresponse)

        End With

        PostToSendMoney(ppusername, pppassword, dollaramount, centamount, recipient, note)

    End Sub
    Protected Sub PostToSendMoney(ByVal ppusername As String, ByVal pppassword As String, ByVal dollaramount As String, ByVal centamount As String, ByVal recipient As String, ByVal note As String)
        'This sub posts all the information needed to send money through paypal.
        Dim url = "https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_flow&SESSION=" & PostLoginSessionData.session & "&dispatch=" & LoginSessionData.dispatch
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        With request
            .Referer = "https://mobile.paypal.com/us/cgi-bin/webscr?cmd=_flow&SESSION=" & LoginSessionData.session & "&dispatch=" & LoginSessionData.dispatch
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "POST"
            Dim sb As New StringBuilder
            sb.Append("CONTEXT=" & PostLoginSessionData.context)
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

            PostSendMoneySessionData = ParseFinalSendMoneyPage(dataresponse)
        End With

        CompleteSendingMoney()

    End Sub
    Protected Sub CompleteSendingMoney()
        'Press the "Send Money" button to complete the payment.
        Dim url = "https://mobile.paypal.com/us/cgi-bin/wapapp?cmd=_flow&SESSION=" & PostSendMoneySessionData.session & "&dispatch=" & LoginSessionData.dispatch
        Dim request As HttpWebRequest = HttpWebRequest.Create(url)
        With request
            .Referer = "https://mobile.paypal.com/us/cgi-bin/webscr?cmd=_flow&SESSION=" & PostLoginSessionData.session & "&dispatch=" & LoginSessionData.dispatch
            .UserAgent = "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.75 Safari/535.7"
            .KeepAlive = True
            .CookieContainer = C
            .Method = "POST"
            Dim sb As New StringBuilder
            sb.Append("CONTEXT=" & PostSendMoneySessionData.context)
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


    End Sub
#End Region

#Region "Login To Normal Paypal And Do Something"

#Region "Login To Paypal"
    Sub Login(ByVal ppusername As String, ByVal pppassword As String, DoWhatAfter As String)
        LoadLoginPage(ppusername, pppassword, DoWhatAfter)
    End Sub
    Protected Sub LoadLoginPage(ByVal ppusername As String, ByVal pppassword As String, DoWhatAfter As String)
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
    Protected Sub PostLoginPage(ByVal ppusername As String, ByVal pppassword As String, DoWhatAfter As String)
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

            Dim r1 As New Regex("conds, <a href=""[A-Za-z0-9=_+-:?%&$#;@]*")
            Dim matches1 As MatchCollection = Regex.Matches(dataresponse, r1.ToString, RegexOptions.Multiline)
            Dim FinalURL = matches1.Item(0).ToString.Replace("conds, <a href=""", "")
            GetMainLoggedInPage(FinalURL, DoWhatAfter)

        End With
    End Sub
    Protected Sub GetMainLoggedInPage(URL As String, DoWhatAfter As String)

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

    End Sub
#End Region
#Region "Grab All Subscribers Ever"
    Protected Sub CollectSessionInfo(dataresponse As String)
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
    End Sub
    Protected Function FindAllEverSubscribers(source As String)
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
    Protected Function FindPagesNeeded(AllEverSubs As Integer)
        Dim pages As Integer = Math.Ceiling(AllEverSubs / 25)
        RaiseEvent Stats("Thread {0} Found Number of Pages to be " & pages & vbNewLine)
        Return pages
    End Function
    Protected Sub FindInformation2(dataresponse As String)
        Dim source = System.Net.WebUtility.HtmlDecode(dataresponse)
        Dim CompleteRegex2 As New Regex(">(?<Name>[^<]*)<[^<]*<[^<]*<[^<]*<[^<]*Email:[^<]*<[^>]*>(?<email>[^<]*)<[^<]*<[^<]*<[^>]*>Customer\sID:[^<]*<[^>]*>(?<CustomerID>[^<]*)<[^<]*<[^<]*<[^>]*>Description:[^<]*<[^<]*<[a-zA-Z0-9&#;]*/p><[^<]*<[^<]*<[^<]*<[^<]*<[^b]*balloonControl[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<DateStart>[^<]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z0-9]*)", RegexOptions.ExplicitCapture)
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

    End Sub
    Protected Sub FindInformation(dataresponse As String)

        Dim source = System.Net.WebUtility.HtmlDecode(dataresponse)

        Dim CompleteRegex As New Regex(">(?<Name>[^<]*)<[^<]*<[^<]*<[^<]*<[^<]*Email:[^<]*<[^>]*>(?<email>[^<]*)<[^<]*<[^<]*<[^>]*>Customer\sID:[^<]*<[^>]*>(?<CustomerID>[^<]*)<[^<]*<[^<]*<[^>]*>Description:[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^<]*<[^>]*>(?<DateStart>[a-zA-Z0-9&#;]*)<[^<]*<[^<]*<[^<]*<[^>]*>(?<Cost>[^<]*)<[^<]*<[^>]*>(?<Active>[a-zA-Z]*)", RegexOptions.ExplicitCapture)
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
    Next
    End Sub
    Protected Sub GetSubscribersPage()
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
            Dim allsubs As Integer = FindAllEverSubscribers(dataresponse)
            PagesInSubscribers = FindPagesNeeded(allsubs)

            CollectSessionInfo(dataresponse)
        End With
        PagesGoneThrough += 1
        GetAllSubscribersNextPage()
    End Sub
    Protected Sub GetAllSubscribersNextPage()
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
    Protected Sub GrabHistoryPage()
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
    Protected Sub GrabDetails(name As String, url As String)
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



    







End Class
