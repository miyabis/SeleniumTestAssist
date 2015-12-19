
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support.PageObjects
Imports OpenQA.Selenium.Support.UI

''' <summary>
''' Seleniumアクション
''' </summary>
Public MustInherit Class SeleniumAction

#Region " Declare "

    Private _driver As IWebDriver
    Private _baseUrl As String
    Private _driverWait As WebDriverWait

#End Region

#Region " コンストラクタ "

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="driver">WebDriverを指定</param>
    Public Sub New(ByVal driver As IWebDriver)
        Me.New(driver, String.Empty)
    End Sub

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="driver">WebDriverを指定</param>
    ''' <param name="baseUrl">ベースとなるURLを指定</param>
    Public Sub New(ByVal driver As IWebDriver, ByVal baseUrl As String)
        _driver = driver
        _baseUrl = baseUrl
        If Not baseUrl.EndsWith("/") Then
            _baseUrl &= "/"
        End If
        Dim timeout As TimeSpan = New TimeSpan(0, 0, 10)
        _driverWait = New WebDriverWait(_driver, timeout)
    End Sub

#End Region

#Region " Property "

    ''' <summary>
    ''' WebDriver
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Driver() As IWebDriver
        Get
            Return _driver
        End Get
    End Property

    ''' <summary>
    ''' ベースURL
    ''' </summary>
    ''' <returns></returns>
    Friend Property BaseUrl As String
        Get
            Return _baseUrl
        End Get
        Set(value As String)
            _baseUrl = value
        End Set
    End Property

    Friend ReadOnly Property driverWait As WebDriverWait
        Get
            Dim timeout As TimeSpan = New TimeSpan(0, 0, 10)
            Return New WebDriverWait(_driver, timeout)
        End Get
    End Property

#End Region

#Region " Actions "

    ''' <summary>
    ''' テキストボックスなどに文字を入力(タイプ)する
    ''' </summary>
    ''' <param name="id">id属性</param>
    ''' <param name="value">入力値</param>
    Public Sub Typing(ByVal id As String, ByVal value As String)
        Typing(Driver.FindElement(By.Id(id)), value)
    End Sub

    ''' <summary>
    ''' テキストボックスなどに文字を入力(タイプ)する
    ''' </summary>
    ''' <param name="element">element</param>
    ''' <param name="value">入力値</param>
    Public Sub Typing(ByVal element As IWebElement, ByVal value As String)
        element.Clear()
        element.SendKeys(value)
    End Sub

    ''' <summary>
    ''' チェックボックスのチェック有無を入力する
    ''' </summary>
    ''' <param name="id">id属性</param>
    ''' <param name="value">入力値</param>
    Public Sub Check(ByVal id As String, ByVal value As Boolean)
        If value AndAlso Driver.FindElement(By.Id(id)).Selected Then
            ' 既にTrueです
            Return
        End If
        If Not value AndAlso Not Driver.FindElement(By.Id(id)).Selected Then
            ' 既にFalseです
            Return
        End If
        Click(id)
    End Sub

    ''' <summary>
    ''' チェックボックスのチェック有無を入力する
    ''' </summary>
    ''' <param name="element">element</param>
    ''' <param name="value">入力値</param>
    Public Sub Check(ByVal element As IWebElement, ByVal value As Boolean)
        If value AndAlso element.Selected Then
            ' 既にTrueです
            Return
        End If
        If Not value AndAlso Not element.Selected Then
            ' 既にFalseです
            Return
        End If
        Click(element)
    End Sub

    ''' <summary>
    ''' 値の取得
    ''' </summary>
    ''' <param name="element"></param>
    ''' <returns></returns>
    Public Function GetValue(ByVal element As IWebElement) As String
        Return element.GetAttribute("Value")
    End Function

    ''' <summary>
    ''' ボタンやリンクをクリックする
    ''' </summary>
    ''' <param name="id">id属性</param>
    Public Sub Click(ByVal id As String)
        Driver.FindElement(By.Id(id)).Click()
    End Sub

    ''' <summary>
    ''' ボタンやリンクをクリックする
    ''' </summary>
    ''' <param name="element">element</param>
    Public Sub Click(ByVal element As IWebElement)
        element.Click()
    End Sub

    ''' <summary>
    ''' select要素(コンボボックス)から選択肢を選ぶ
    ''' </summary>
    ''' <param name="id">id属性</param>
    ''' <param name="value">ラベル文字列</param>
    Public Sub SelectByText(ByVal id As String, ByVal value As String)
        Dim element As IWebElement
        Dim selectElement As Support.UI.SelectElement
        element = Driver.FindElement(By.Id(id))
        selectElement = New Support.UI.SelectElement(element)
        selectElement.SelectByText(value)
    End Sub

    ''' <summary>
    ''' select要素(コンボボックス)から選択肢を選ぶ
    ''' </summary>
    ''' <param name="id">id属性</param>
    ''' <param name="value">value属性値</param>
    Public Sub SelectByValue(ByVal id As String, ByVal value As String)
        Dim element As IWebElement
        Dim selectElement As Support.UI.SelectElement
        element = Driver.FindElement(By.Id(id))
        selectElement = New Support.UI.SelectElement(element)
        selectElement.SelectByValue(value)
    End Sub

    ''' <summary>
    ''' select要素(コンボボックス)から選択肢を選ぶ
    ''' </summary>
    ''' <param name="id">id属性</param>
    ''' <param name="value">インデックス番号</param>
    Public Sub SelectByIndex(ByVal id As String, ByVal value As Integer)
        Dim element As IWebElement
        Dim selectElement As Support.UI.SelectElement
        element = Driver.FindElement(By.Id(id))
        selectElement = New Support.UI.SelectElement(element)
        selectElement.SelectByIndex(value)
    End Sub

    ''' <summary>
    ''' FindElement をデフォルト１０秒まで取得できるまで待つ
    ''' </summary>
    ''' <param name="by"></param>
    ''' <returns></returns>
    Public Function FindElementWaitUntil(ByVal by As By) As IWebElement
        Return driverWait.Until(
            Function(d)
                Dim el As IWebElement
                el = d.FindElement(by)
                If el.Displayed And el.Enabled Then
                    Return el
                End If
                Return Nothing
            End Function)
    End Function

    Public Function createPage(Of T)() As T
        Dim page As T = PageFactory.InitElements(Of T)(Driver)
        Dim pageAction As Object = page

        If TypeOf pageAction Is AbstractSeleniumTest Then
            CType(pageAction, SeleniumAction).BaseUrl = _baseUrl
        End If
        Return page
    End Function

#End Region

End Class
