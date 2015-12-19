
Imports MiYABiS.SeleniumTestAssist

<TestClass(),
    DeploymentItem("IEDriverServer.exe")>
Public Class UnitTest1
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
        Try
            SeleniumCleanup()
        Finally
            IISExpressManager.Stop()
        End Try
    End Sub

    ''' <summary>
    ''' 各テストを実行する前にコードを実行するには、TestInitialize を使用
    ''' </summary>
    ''' <remarks></remarks>
    <TestInitialize()>
    Public Overrides Sub TestInitialize()
        MyBase.TestInitialize()
    End Sub

    ''' <summary>
    ''' 各テストを実行した後にコードを実行するには、TestCleanup を使用
    ''' </summary>
    ''' <remarks></remarks>
    <TestCleanup()>
    Public Overrides Sub TestCleanup()
        MyBase.TestCleanup("最後")
    End Sub

#End Region

    <TestMethod(),
     Description("テストが表示されること"),
     TestCategory("表示系")>
    Public Sub TestMethod1()
        IEInitialize()

        Open("Default.aspx", 1000, 1000)

        Dim page As DefaultPage
        page = createPage(Of DefaultPage)()

        page.HogeAssert("テスト")

        getScreenshot()

        page.BtnTest()

        getScreenshot()

        page.TestAssert("btnTest Click!")

        Dim page2 As AboutPage
        page2 = page.About()

        page2.H2Assert()
    End Sub

    'DataSource("System.Data.Odbc", "Dsn=Excel Files;Driver={Microsoft Excel Driver (*.xls)};dbq=|DataDirectory|\\ExcelTestData.xls", "Login$", DataAccessMethod.Sequential)
    <TestMethod(),
     Description("ログインエラーとなること"),
     TestCategory("エラー系"),
     DeploymentItem("TestData.csv"),
     DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "TestData.csv", "TestData#csv", DataAccessMethod.Sequential)
     >
    Public Sub TestMethod2()
        IEInitialize()

        Open("/Account/Login", 1000, 1000)

        Dim page As LoginPage
        page = createPage(Of LoginPage)()

        getScreenshot("オープン直後")

        page.Email(Me.TestContext.DataRow("ID"))
        page.Password(Me.TestContext.DataRow("Pass"))
        page.RememberMe(True)

        getScreenshot("入力後", "ノートがこんな感じで表示されます。")

        page.LogIn()
    End Sub

    <TestMethod(),
     Description("Demoページテスト"),
     TestCategory("正常系")>
    Public Sub TestMethod3()
        IEInitialize()
        'FirefoxInitialize()
        'ChromeInitialize()
        'EdgeInitialize("C:\Program Files (x86)\Microsoft Web Driver\")

        Open("/Demo")

        Dim page As DemoPage
        page = createPage(Of DemoPage)()

        page.TextBox1("test")
        page.CheckBox1(True)
        page.DropDownList1("demo2")
    End Sub

End Class
