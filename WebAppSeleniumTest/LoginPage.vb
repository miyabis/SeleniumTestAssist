
Imports OpenQA.Selenium
Imports MiYABiS.SeleniumTestAssist

Public Class LoginPage
    Inherits SeleniumAction

    Public Overrides ReadOnly Property MyPageName As String
        Get
            Return "/Account/Login"
        End Get
    End Property

    Public Sub New(ByVal driver As IWebDriver, ByVal baseUrl As String)
        MyBase.New(driver, baseUrl)
    End Sub

    Public Sub Email(ByVal value As String)
        Driver.FindElement(By.Id("MainContent_Email")).Clear()
        Driver.FindElement(By.Id("MainContent_Email")).SendKeys(value)
    End Sub

    Public Sub Password(ByVal value As String)
        Driver.FindElement(By.Id("MainContent_Password")).Clear()
        Driver.FindElement(By.Id("MainContent_Password")).SendKeys(value)
    End Sub

    Public Sub RememberMe(ByVal value As Boolean)
        Driver.FindElement(By.Id("MainContent_RememberMe")).Click()
    End Sub

End Class
