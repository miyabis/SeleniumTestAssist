
Imports OpenQA.Selenium
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

#Region " MustOverride "

    ''' <summary>
    ''' 対象となるページ又はコマンド
    ''' </summary>
    ''' <returns></returns>
    Public MustOverride ReadOnly Property MyPageName As String

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
    Public ReadOnly Property BaseUrl As String
        Get
            Return _baseUrl
        End Get
    End Property

#End Region

#Region " Actions "

    ''' <summary>
    ''' 指定したURLを開く
    ''' </summary>
    ''' <remarks>Aspx又はMVCのコマンドを開く</remarks>
    Public Overloads Sub Open()
        Driver.Navigate().GoToUrl(String.Format("{0}{1}", BaseUrl, MyPageName))
    End Sub

    ''' <summary>
    ''' 指定したURLを開く
    ''' </summary>
    ''' <param name="width">ウィンドウの幅を指定</param>
    ''' <param name="height">ウィンドウの高さを指定</param>
    ''' <remarks>Aspx又はMVCのコマンドを開く</remarks>
    Public Overloads Sub Open(ByVal width As Integer, ByVal height As Integer)
        Driver.Manage.Window.Size = New Drawing.Size(width, height)
        Me.Open()
    End Sub

    ''' <summary>
    ''' テキストボックスなどに文字を入力(タイプ)する
    ''' </summary>
    ''' <param name="id">id属性</param>
    ''' <param name="value">入力値</param>
    Public Sub Typing(ByVal id As String, ByVal value As String)
        Driver.FindElement(By.Id(id)).Clear()
        Driver.FindElement(By.Id(id)).SendKeys(value)
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
    ''' ボタンやリンクをクリックする
    ''' </summary>
    ''' <param name="id">id属性</param>
    Public Sub Click(ByVal id As String)
        Driver.FindElement(By.Id(id)).Click()
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

    'TODO:
    '''' <summary>
    '''' FindElement をデフォルト１０秒まで取得できるまで待つ
    '''' </summary>
    '''' <param name="by"></param>
    '''' <returns></returns>
    'Public Function FindElementWaitUntil(ByVal by As By) As IWebElement
    '    Return _driverWait.Until(
    '        Function(d)
    '            Dim el As IWebElement
    '            el = d.FindElement(by)
    '            If el.Displayed And el.Enabled Then
    '                Return el
    '            End If
    '            Return Nothing
    '        End Function)
    'End Function

#End Region

End Class
