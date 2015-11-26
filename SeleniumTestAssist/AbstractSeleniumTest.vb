
Imports System.Text
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports OpenQA.Selenium
Imports OpenQA.Selenium.Remote
Imports OpenQA.Selenium.IE
Imports OpenQA.Selenium.Support.UI
Imports Selenium
Imports OpenQA.Selenium.Firefox

''' <summary>
''' Selemium2 を使ったブラウザテスト用抽象クラス
''' </summary>
''' <remarks>
''' http://docs.seleniumhq.org/ <br/>
''' </remarks>
Public MustInherit Class AbstractSeleniumTest

#Region " Declare "

    Const INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS As String = "ignoreProtectedModeSettings"
    Const ENABLE_ELEMENT_CACHE_CLEANUP As String = "enableElementCacheCleanup"
    Const IE_ENSURE_CLEAN_SESSION As String = "ie.ensureCleanSession"

    Protected capabilities As ICapabilities
    Protected driver As IWebDriver
    Protected javaScriptExecutor As IJavaScriptExecutor
    Protected driverWait As WebDriverWait
    Protected timeout As TimeSpan
    Protected verificationErrors As StringBuilder
    Protected acceptNextAlert As Boolean = True
    Protected screenshotCount As Integer

    Protected testContextInstance As TestContext

    Private Shared _baseUrl As String

#Region " Logging For Log4net "
    ''' <summary>Logging For Log4net</summary>
    Private Shared ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(String.Empty)
#End Region
#End Region

#Region " Property "

    '''<summary>
    '''現在のテストの実行についての情報および機能を
    '''提供するテスト コンテキストを取得または設定します。
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(ByVal value As TestContext)
            testContextInstance = value
        End Set
    End Property

#End Region

#Region " Methods "

    ''' <summary>
    ''' クラスの最初のテストを実行する前にコードを実行するには、ClassInitialize 属性を持つメソッドで使用してください。
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub SeleniumInitialize(ByVal baseUrl As String)
        _baseUrl = baseUrl
    End Sub

    ''' <summary>
    ''' クラスのすべてのテストを実行した後にコードを実行するには、ClassCleanup 属性を持つメソッドで使用してください。
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub SeleniumCleanup()
    End Sub

    ''' <summary>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' 必要であればオーバーライドしてください。
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub TestInitialize()
    End Sub

    ''' <summary>
    ''' 各テストを実行した後にコードを実行するには、TestCleanup 属性を持つメソッドで使用してください。
    ''' 必要であればオーバーライドしてください。
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub TestCleanup()
        Try
            If driver IsNot Nothing Then
                getScreenshot()
                driver.Quit()
            End If
        Catch ex As Exception
            ' Ignore errors if unable to close the browser
        Finally
        End Try
        Assert.AreEqual("", verificationErrors.ToString())
    End Sub

#Region " IE "

    ''' <summary>
    ''' ローカルの Selenium で IE 実行する時の初期化
    ''' </summary>
    ''' <param name="ieDriverServerPath"></param>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Sub IEInitialize(Optional ByVal ieDriverServerPath As String = Nothing)
        Dim ieDriverServer As String = ieDriverServerPath
        Dim opt As InternetExplorerOptions = New InternetExplorerOptions()

        If String.IsNullOrEmpty(ieDriverServer) Then
            ieDriverServer = AppDomain.CurrentDomain.BaseDirectory
            ieDriverServer = My.Application.Info.DirectoryPath
        End If
        opt.IntroduceInstabilityByIgnoringProtectedModeSettings = True
        driver = New InternetExplorerDriver(ieDriverServer, opt)
        capabilities = DirectCast(driver, InternetExplorerDriver).Capabilities

        Dim callingMethod = New System.Diagnostics.StackTrace(1, False).GetFrame(0).GetMethod()
        Dim attrs() As Attribute = callingMethod.GetCustomAttributes(GetType(DescriptionAttribute), False)
        Dim description As DescriptionAttribute = Nothing
        For Each attr As Attribute In attrs
            If TypeOf attr Is DescriptionAttribute Then
                description = attr
            End If
        Next
        If description IsNot Nothing Then
            Me.TestContext.WriteLine(description.Description)
        End If

        _testInitialize()
    End Sub

    ''' <summary>
    ''' SeleniumRC で IE 実行する時の初期化
    ''' </summary>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Sub IERemoteInitialize(ByVal seleniumURL As String, Optional ByVal version As String = Nothing)
        Dim ieCapability As DesiredCapabilities = DesiredCapabilities.InternetExplorer()
        ieCapability.SetCapability(INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, True)
        If Not String.IsNullOrEmpty(version) Then
            ieCapability.SetCapability("version", version)
        End If
        'ieCapability.SetCapability("logFile", "E:\Temp\selenium.log")
        'ieCapability.SetCapability("logLevel", "TRACE")

        driver = New RemoteWebDriver(New Uri(seleniumURL), ieCapability)
        capabilities = ieCapability

        _testInitialize()
    End Sub

#End Region
#Region " Edge "

    ''' <summary>
    ''' ローカルの Selenium で Edge 実行する時の初期化
    ''' </summary>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Sub EdgeInitialize(Optional ByVal ieDriverServerPath As String = Nothing)
        Dim ieDriverServer As String = ieDriverServerPath

        If ieDriverServer Is Nothing Then
            driver = New Edge.EdgeDriver()
        Else
            driver = New Edge.EdgeDriver(ieDriverServer)
        End If
        capabilities = DirectCast(driver, Edge.EdgeDriver).Capabilities

        _testInitialize()
    End Sub

    ''' <summary>
    ''' SeleniumRC で Edge 実行する時の初期化
    ''' </summary>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Sub EdgeRemoteInitialize(ByVal seleniumURL As String, Optional ByVal version As String = Nothing)
        Dim ieCapability As DesiredCapabilities = DesiredCapabilities.Edge()
        If Not String.IsNullOrEmpty(version) Then
            ieCapability.SetCapability("version", version)
        End If

        driver = New RemoteWebDriver(New Uri(seleniumURL), ieCapability)
        capabilities = ieCapability

        _testInitialize()
    End Sub

#End Region
#Region " Firefox "

    ''' <summary>
    ''' ローカルの Selenium で Firefox 実行する時の初期化
    ''' </summary>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Overloads Sub FirefoxInitialize()
        driver = New Firefox.FirefoxDriver()
        capabilities = DirectCast(driver, Firefox.FirefoxDriver).Capabilities

        _testInitialize()
    End Sub

    ''' <summary>
    ''' SeleniumRC で Firefox 実行する時の初期化
    ''' </summary>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Overloads Sub FirefoxRemoteInitialize(ByVal seleniumURL As String, Optional ByVal version As String = Nothing)
        Dim ieCapability As DesiredCapabilities = DesiredCapabilities.Firefox()
        If Not String.IsNullOrEmpty(version) Then
            ieCapability.SetCapability("version", version)
        End If

        driver = New RemoteWebDriver(New Uri(seleniumURL), ieCapability)
        capabilities = ieCapability

        _testInitialize()
    End Sub

    ''' <summary>
    ''' ローカルの Selenium で Firefox 実行する時の初期化
    ''' </summary>
    ''' <param name="proxyHost">プロキシのホスト</param>
    ''' <param name="proxyPort">プロキシのポート</param>
    ''' <param name="autoAuthHosts">プロキシなしで接続するホスト名のカンマ区切り</param>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Overloads Sub FirefoxInitialize(ByVal proxyHost As String, ByVal proxyPort As String, Optional ByVal autoAuthHosts As String = Nothing)
        Dim proxy As New OpenQA.Selenium.Proxy
        proxy.HttpProxy = proxyHost & ":" & proxyPort

        Dim profile As New FirefoxProfile()
        profile.SetProxyPreferences(proxy)
        If Not String.IsNullOrEmpty(autoAuthHosts) Then
            profile.SetPreference("network.proxy.no_proxies_on", autoAuthHosts)
            profile.SetPreference("network.negotiate-auth.trusted-uris", autoAuthHosts)
            'profile.SetPreference("network.automatic-ntlm-auth.trusted-uris", autoAuthHosts)
            'profile.SetPreference("network.negotiate-auth.delegation-uris", autoAuthHosts)
        End If

        driver = New Firefox.FirefoxDriver(profile)
        capabilities = DirectCast(driver, Firefox.FirefoxDriver).Capabilities

        _testInitialize()
    End Sub

    ''' <summary>
    ''' SeleniumRC で Firefox 実行する時の初期化
    ''' </summary>
    ''' <param name="seleniumURL">SeleniumRC の URL</param>
    ''' <param name="proxyHost">プロキシのホスト</param>
    ''' <param name="proxyPort">プロキシのポート</param>
    ''' <param name="autoAuthHosts">プロキシなしで接続するホスト名のカンマ区切り</param>
    ''' <param name="version">実行するブラウザのバージョン</param>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Overloads Sub FirefoxRemoteInitialize(ByVal seleniumURL As String, ByVal proxyHost As String, ByVal proxyPort As String, Optional ByVal autoAuthHosts As String = Nothing, Optional ByVal version As String = Nothing)
        Dim ieCapability As DesiredCapabilities = DesiredCapabilities.Firefox()
        If Not String.IsNullOrEmpty(version) Then
            ieCapability.SetCapability("version", version)
        End If

        Dim proxy As New OpenQA.Selenium.Proxy
        proxy.HttpProxy = proxyHost & ":" & proxyPort

        Dim profile As New FirefoxProfile()
        profile.SetProxyPreferences(proxy)
        If Not String.IsNullOrEmpty(autoAuthHosts) Then
            profile.SetPreference("network.proxy.no_proxies_on", autoAuthHosts)
            profile.SetPreference("network.negotiate-auth.trusted-uris", autoAuthHosts)
            'profile.SetPreference("network.automatic-ntlm-auth.trusted-uris", autoAuthHosts)
            'profile.SetPreference("network.negotiate-auth.delegation-uris", autoAuthHosts)
        End If
        ieCapability.SetCapability(FirefoxDriver.ProfileCapabilityName, profile.ToBase64String())

        driver = New RemoteWebDriver(New Uri(seleniumURL), ieCapability)

        capabilities = ieCapability

        _testInitialize()
    End Sub

#End Region
#Region " Chrome "

    ''' <summary>
    ''' ローカルの Selenium で Chrome 実行する時の初期化
    ''' </summary>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Sub ChromeInitialize()
        driver = New Chrome.ChromeDriver()
        capabilities = DirectCast(driver, Chrome.ChromeDriver).Capabilities

        _testInitialize()
    End Sub

    ''' <summary>
    ''' SeleniumRC で Chrome 実行する時の初期化
    ''' </summary>
    ''' <remarks>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize 属性を持つメソッドで使用してください。
    ''' </remarks>
    Protected Sub ChromeRemoteInitialize(ByVal seleniumURL As String, Optional ByVal version As String = Nothing)
        Dim ieCapability As DesiredCapabilities = DesiredCapabilities.Chrome()
        If Not String.IsNullOrEmpty(version) Then
            ieCapability.SetCapability("version", version)
        End If

        driver = New RemoteWebDriver(New Uri(seleniumURL), ieCapability)
        capabilities = ieCapability

        _testInitialize()
    End Sub

#End Region

    Protected Function isElementPresent(by As By) As Boolean
        Try
            driver.FindElement(by)
            Return True
        Catch ex As NoSuchElementException
            Return False
        End Try
    End Function

    Protected Function isAlertPresent() As Boolean
        Try
            driver.SwitchTo().Alert()
            Return True
        Catch ex As NoAlertPresentException
            Return False
        End Try
    End Function

    Protected Function closeAlertAndGetItsText() As String
        Try
            Dim wait As New WebDriverWait(driver, TimeSpan.FromSeconds(10))
            wait.Until(Function(d)
                           Return driver.SwitchTo().Alert() IsNot Nothing
                       End Function)
            Dim alert As IAlert = driver.SwitchTo().Alert()
            Dim alertText As String = alert.Text
            If acceptNextAlert Then
                alert.Accept()
            Else
                alert.Dismiss()
            End If
            Return alertText
        Finally
            acceptNextAlert = True
        End Try
    End Function

    Protected Sub getScreenshot()
        Dim savePath As String = Path.Combine(Me.TestContext.ResultsDirectory, Me.TestContext.FullyQualifiedTestClassName.Replace(".", "\"))
        If Not Directory.Exists(savePath) Then
            Directory.CreateDirectory(savePath)
        End If
        Dim filename As String = String.Format("{0}_{1}.png", Me.TestContext.TestName, screenshotCount)
        Dim fullPath As String = Path.Combine(savePath, filename)
        CType(driver, ITakesScreenshot).GetScreenshot().SaveAsFile(fullPath, System.Drawing.Imaging.ImageFormat.Png)
        Me.TestContext.AddResultFile(fullPath)
        screenshotCount += 1
    End Sub

    Protected Sub sleep(Optional ByVal value As Integer = 300)
        System.Threading.Thread.Sleep(value)
    End Sub

    Protected Function createPage(Of T)() As T
        Return CType(Activator.CreateInstance(GetType(T), New Object() {driver, _baseUrl}), T)
    End Function

    Private Sub _testInitialize()
        timeout = New TimeSpan(0, 0, 10)
        driverWait = New WebDriverWait(driver, timeout)
        driver.Manage.Timeouts.ImplicitlyWait(timeout)

        javaScriptExecutor = DirectCast(driver, IJavaScriptExecutor)

        verificationErrors = New StringBuilder()
        screenshotCount = 0

        Me.TestContext.WriteLine("WebDriver：{0}", capabilities.ToString)
        _mylog.DebugFormat("WebDriver：{0}", capabilities.ToString)
    End Sub

#End Region

End Class
