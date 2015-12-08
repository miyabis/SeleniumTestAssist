
Imports MiYABiS.SeleniumTestAssist

<TestClass()> Public Class UnitTest1
    Inherits AbstractSeleniumTest

#Region " Declare "

    Private Const _PORT As Integer = 50744

    Private Shared _baseUrl As String = String.Format("http://localhost:{0}/", _PORT)

#Region " Logging For Log4net "
    ''' <summary>Logging For Log4net</summary>
    Private Shared ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(String.Empty)
#End Region
#End Region

#Region "追加のテスト属性"

    ''' <summary>
    ''' クラスの最初のテストを実行する前にコードを実行するには、ClassInitialize を使用
    ''' </summary>
    ''' <param name="testContext"></param>
    ''' <remarks></remarks>
    <ClassInitialize()>
    Public Shared Sub ClassInitialize(ByVal testContext As TestContext)
        IISExpressManager.ProjectName = "WebApp"
        'IISExpressManager.VPath = "/Test"
        IISExpressManager.Port = _PORT
        'IISExpressManager.Clr = "v4.0"
        'IISExpressManager.Ntlm = False
        IISExpressManager.Start()

        SeleniumInitialize(_baseUrl)
    End Sub

    ''' <summary>
    ''' クラスのすべてのテストを実行した後にコードを実行するには、ClassCleanup を使用
    ''' </summary>
    ''' <remarks></remarks>
    <ClassCleanup()>
    Public Shared Sub ClassCleanup()
        SeleniumCleanup()

        IISExpressManager.Stop()
    End Sub

    ''' <summary>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize を使用
    ''' </summary>
    ''' <remarks></remarks>
    <TestInitialize()>
    Public Overrides Sub TestInitialize()
    End Sub

    ''' <summary>
    ''' 各テストを実行した後にコードを実行するには、TestCleanup を使用
    ''' </summary>
    ''' <remarks></remarks>
    <TestCleanup()>
    Public Overrides Sub TestCleanup()
        MyBase.TestCleanup()
    End Sub

#End Region

    <TestMethod(),
     Description("テストが表示されること"),
     TestCategory("表示系")>
    Public Sub TestMethod1()
        IEInitialize()

        Dim page As DefaultPage
        page = createPage(Of DefaultPage)()

        page.Open(1000, 1000)

        page.HogeAssert("テスト")

        page.BtnTest.Click()
        page.TestAssert("btnTest Click!")
    End Sub

    <TestMethod(),
     Description("ログインエラーとなること"),
     TestCategory("エラー系")>
    Public Sub TestMethod2()
        IEInitialize()

        Dim page As LoginPage
        page = createPage(Of LoginPage)()

        page.Open(1000, 1000)

        getScreenshot("オープン直後")

        page.Email("test")
        page.Password("hoge")
        page.RememberMe(True)
    End Sub

    <TestMethod(),
     Description("Demoページテスト"),
     TestCategory("正常系")>
    Public Sub TestMethod3()
        'IEInitialize()
        'FirefoxInitialize()
        'ChromeInitialize()
        EdgeInitialize("C:\Program Files (x86)\Microsoft Web Driver\")

        Dim page As DemoPage
        page = createPage(Of DemoPage)()

        'page.Open(1000, 1000)
        page.Open()

        page.TextBox1("test")
        page.CheckBox1(True)
        page.DropDownList1("demo2")
    End Sub

End Class
