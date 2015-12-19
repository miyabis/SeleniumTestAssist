
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support.PageObjects
Imports MiYABiS.SeleniumTestAssist

Public Class DefaultPage
    Inherits SeleniumAction

    <FindsBy(How:=How.CssSelector, [Using]:=".hoge")>
    Private _hoge As IWebElement

    <FindsBy([Using]:="MainContent_txtTest")>
    Private _test As IWebElement

    <FindsBy([Using]:="MainContent_btnTest")>
    Private _BtnTest As IWebElement

    <FindsBy([Using]:="lnkAbout")>
    Private _aboutLink As IWebElement

    Public Sub New(ByVal driver As IWebDriver)
        MyBase.New(driver)
    End Sub

    Public Sub BtnTest()
        _BtnTest.Click()
    End Sub

    Public Function About() As AboutPage
        Click(_aboutLink)
        Return createPage(Of AboutPage)()
    End Function

    Public Sub HogeAssert(ByVal value As String)
        Assert.AreEqual(value, _hoge.Text)
    End Sub

    Public Sub TestAssert(ByVal value As String)
        Dim element As IWebElement = FindElementWaitUntil(By.Id("MainContent_txtTest"))
        Assert.AreEqual(value, GetValue(_test))
    End Sub

End Class
