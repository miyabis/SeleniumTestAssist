﻿
Imports System.Text
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports OpenQA.Selenium
Imports OpenQA.Selenium.Remote
Imports OpenQA.Selenium.IE
Imports OpenQA.Selenium.Support.UI
Imports OpenQA.Selenium.Firefox
Imports OfficeOpenXml
Imports OpenQA.Selenium.Support.PageObjects

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

    Protected Friend Shared testMethods As IDictionary(Of String, Integer)
    Protected Friend Shared screenshots As IList(Of ScreenshotRow)
    Protected Friend Shared callingAssembly As Reflection.Assembly

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
            If testMethods.ContainsKey(Me.TestContext.TestName) Then
                Dim cnt As Integer
                cnt = testMethods(Me.TestContext.TestName)
                testMethods(Me.TestContext.TestName) = cnt + 1
            Else
                testMethods.Add(Me.TestContext.TestName, 1)
            End If
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
        screenshots = New List(Of ScreenshotRow)
        testMethods = New Dictionary(Of String, Integer)
        callingAssembly = Reflection.Assembly.GetCallingAssembly
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
    Public Overridable Overloads Sub TestCleanup()
        Me.TestCleanup(String.Empty, String.Empty)
    End Sub

    Public Overloads Sub TestCleanup(ByVal title As String, Optional ByVal note As String = "")
        Try
            If driver IsNot Nothing Then
                getScreenshot(title, note)
            End If
        Finally
            If driver IsNot Nothing Then
                driver.Quit()
            End If
        End Try
        Assert.AreEqual("", verificationErrors.ToString())

        outputExcel()

        screenshots = New List(Of ScreenshotRow)
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

        Dim callingMethod = New StackTrace(1, False).GetFrame(0).GetMethod()
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

    ''' <summary>
    ''' 指定したURLを開く
    ''' </summary>
    ''' <remarks>Aspx又はMVCのコマンドを開く</remarks>
    Protected Sub Open(ByVal testPageUrl As String)
        driver.Navigate().GoToUrl(String.Format("{0}{1}", _baseUrl, testPageUrl))
    End Sub

    ''' <summary>
    ''' 指定したURLを開く
    ''' </summary>
    ''' <param name="testPageUrl"></param>
    ''' <param name="width">ウィンドウの幅を指定</param>
    ''' <param name="height">ウィンドウの高さを指定</param>
    ''' <remarks>Aspx又はMVCのコマンドを開く</remarks>
    Protected Sub Open(ByVal testPageUrl As String, ByVal width As Integer, ByVal height As Integer)
        driver.Manage.Window.Size = New System.Drawing.Size(width, height)
        Me.Open(testPageUrl)
    End Sub

    Protected Sub outputExcel()
        Dim fname As New FileInfo(_getExcelEvidenceName())

        Using package As New ExcelPackage(fname)
            Dim ws As ExcelWorksheet = Nothing

            Dim colsMax As Integer = 12
            Dim oneColWidth As Integer = 64

            Dim colIndex As Integer = 3
            Dim rowIndex As Integer = 5
            Dim sheetname As String = String.Empty
            Dim cell As ExcelRange
            Dim imgCount As Integer

            For Each row As ScreenshotRow In screenshots
                If sheetname <> row.TestMethodName Then
                    Dim name As String
                    'Dim lstWs() As ExcelWorksheet
                    'lstWs = package.Workbook.Worksheets.Select(Of ExcelWorksheet)(
                    'Function(sheet)
                    '    If sheet.Name.StartsWith(row.TestMethodName) Then
                    '        Return sheet
                    '    End If
                    '    Return Nothing
                    'End Function).ToArray
                    'name = row.TestMethodName
                    'If Not lstWs.Count.Equals(0) Then
                    '    name &= "_" & (lstWs.Count + 1)
                    'End If
                    name = String.Format("{0}_{1:000}", row.TestMethodName, testMethods(row.TestMethodName))
                    ws = package.Workbook.Worksheets.Add(name)
                    ws.Cells.Style.Font.SetFromFont(New System.Drawing.Font("Meiryo UI", 10, System.Drawing.FontStyle.Regular))
                    cell = ws.Cells(2, 2)
                    cell.Value = row.TestMethodName
                    cell.Style.Font.Size = 16
                    cell = ws.Cells(2, 2, 2, 15)
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                    cell.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.RoyalBlue)
                    sheetname = row.TestMethodName
                    rowIndex = 5
                    imgCount = 1

                    Dim asm As System.Reflection.Assembly
                    'asm = System.Reflection.Assembly.GetCallingAssembly
                    asm = callingAssembly
                    Dim typ As Type
                    typ = asm.GetType(row.TestClassName)
                    Dim method As System.Reflection.MethodInfo
                    method = typ.GetMethod(row.TestMethodName)
                    Dim attrs() As Attribute
                    attrs = method.GetCustomAttributes(GetType(DescriptionAttribute), False)
                    If Not attrs.Count.Equals(0) Then
                        cell = ws.Cells(2, 15)
                        cell.Value = CType(attrs(0), DescriptionAttribute).Description
                        cell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right
                    End If
                End If

                Dim img As System.Drawing.Image
                img = System.Drawing.Image.FromFile(row.FullPath)
                Dim picture = ws.Drawings.AddPicture(row.Filename, img)
                picture.SetPosition(rowIndex, 0, colIndex, 0)
                Dim height As Integer = picture.Image.Height
                If picture.Image.Size.Width > (colsMax * oneColWidth) Then
                    Dim val As Integer
                    val = ((colsMax * oneColWidth) / picture.Image.Width) * 100 - 1
                    picture.SetSize(val)
                    height = height * (val / 100)
                End If

                cell = ws.Cells(rowIndex - 1, colIndex)
                cell.Value = String.Format("{0}. {1}", imgCount, row.Title)
                cell.Style.Font.Size = 14

                cell = ws.Cells(rowIndex - 1, 15)
                cell.Value = row.Note
                cell.Style.Font.Size = 11
                cell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                cell = ws.Cells(rowIndex - 1, 3, rowIndex - 1, 15)
                cell.Style.Font.Color.SetColor(System.Drawing.Color.White)
                cell.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue)

                Dim rowsCount As Integer = Math.Ceiling(height / 20)
                rowIndex += rowsCount + 4

                cell = ws.Cells(rowIndex - 3, 4)
                cell.Value = row.WindowTitle
                cell.Style.Font.Size = 9
                cell.Style.Font.Color.SetColor(System.Drawing.Color.Gray)

                cell = ws.Cells(rowIndex - 3, 15)
                cell.Value = row.Filename
                cell.Style.Font.Size = 9
                cell.Style.Font.Color.SetColor(System.Drawing.Color.Silver)
                cell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                imgCount += 1
            Next

            package.Save()
        End Using

        Me.TestContext.AddResultFile(fname.FullName)
    End Sub

    'TODO:
    '''' <summary>
    '''' FindElement をデフォルト１０秒まで取得できるまで待つ
    '''' </summary>
    '''' <param name="by"></param>
    '''' <returns></returns>
    'Protected Function FindElementWaitUntil(ByVal by As By) As IWebElement
    '    Return driverWait.Until(
    '        Function(d)
    '            Dim el As IWebElement
    '            el = d.FindElement(by)
    '            If el.Displayed And el.Enabled Then
    '                Return el
    '            End If
    '            Return Nothing
    '        End Function)
    'End Function

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

    Protected Function getSavePath()
        Dim savePath As String = Path.Combine(Me.TestContext.ResultsDirectory, Me.TestContext.FullyQualifiedTestClassName.Replace(".", "\"))
        If Not Directory.Exists(savePath) Then
            Directory.CreateDirectory(savePath)
        End If
        Return savePath
    End Function

    Protected Overloads Sub getScreenshot()
        getScreenshot(Nothing)
    End Sub

    Protected Overloads Sub getScreenshot(ByVal title As String, Optional ByVal note As String = "")
        Dim savePath As String = getSavePath()
        Dim fn As String = String.Format("{0}_{1:000}_{2:00000}{3}.png", Me.TestContext.TestName, testMethods(Me.TestContext.TestName), screenshotCount, IIf(String.IsNullOrEmpty(title), String.Empty, "_" & title))
        Dim fullPath As String = Path.Combine(savePath, fn)
        CType(driver, ITakesScreenshot).GetScreenshot().SaveAsFile(fullPath, System.Drawing.Imaging.ImageFormat.Png)
        Me.TestContext.AddResultFile(fullPath)

        Dim row As New ScreenshotRow
        row.TestClassName = Me.TestContext.FullyQualifiedTestClassName
        row.TestMethodName = Me.TestContext.TestName
        row.FullPath = fullPath
        row.Title = title
        row.Count = screenshotCount
        row.Note = note
        row.WindowTitle = driver.Title

        screenshots.Add(row)

        screenshotCount += 1
    End Sub

    Protected Sub sleep(Optional ByVal value As Integer = 300)
        Threading.Thread.Sleep(value)
    End Sub

    Protected Function createPage(Of T)() As T
        Dim page As T = PageFactory.InitElements(Of T)(driver)
        Dim pageAction As Object = page

        If TypeOf pageAction Is AbstractSeleniumTest Then
            CType(pageAction, SeleniumAction).BaseUrl = _baseUrl
        End If
        Return page
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

    Private Function _getExcelEvidenceName() As String
        Dim savePath As String = getSavePath()
        Return Path.Combine(savePath, String.Format("{0}.xlsx", Me.GetType().Name))
    End Function

#End Region

End Class



<AttributeUsage(AttributeTargets.Class, AllowMultiple:=False)>
Public Class ExcelEvidenceAttribute
    Inherits Attribute

    Private _filename As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal filename As String)
        _filename = filename
    End Sub

End Class

