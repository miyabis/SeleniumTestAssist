
Imports OpenQA.Selenium
Imports MiYABiS.SeleniumTestAssist

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

    Public Function HogeElement() As IWebElement
        Return Driver.FindElement(By.CssSelector(".hoge"))
    End Function

    Public Sub HogeElementAssert(ByVal value As String)
        Assert.AreEqual(Me.HogeElement().Text, value)
    End Sub

End Class
