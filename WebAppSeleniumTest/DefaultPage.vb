
Imports OpenQA.Selenium
Imports MiYABiS.SeleniumTestAssist
Imports OpenQA.Selenium.Support.UI

Public Class DefaultPage
    Inherits SeleniumAction

    Public Overrides ReadOnly Property MyPageName As String
        Get
            Return "Default.aspx"
        End Get
    End Property

    Public Sub New(ByVal driver As IWebDriver, ByVal baseUrl As String)
        MyBase.New(driver, baseUrl)
    End Sub

    Public ReadOnly Property Hoge() As String
        Get
            Return Driver.FindElement(By.CssSelector(".hoge")).Text
        End Get
    End Property

    Public Property Test() As String
        Get
            Return Driver.FindElement(By.Id("MainContent_txtTest")).GetAttribute("Value")
            'Return FindElementWaitUntil(By.Id("MainContent_txtTest")).GetAttribute("Value")
        End Get
        Set(value As String)
            Typing("MainContent_txtTest", value)
        End Set
    End Property

    Public ReadOnly Property BtnTest() As IWebElement
        Get
            Return Driver.FindElement(By.Id("MainContent_btnTest"))
        End Get
    End Property

    Public Sub HogeAssert(ByVal value As String)
        Assert.AreEqual(value, Me.Hoge)
    End Sub

    Public Sub TestAssert(ByVal value As String)
        Assert.AreEqual(value, Me.Test)
    End Sub

End Class
