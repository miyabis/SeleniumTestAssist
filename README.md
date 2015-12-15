# SeleniumTestAssist

[![Build status](https://ci.appveyor.com/api/projects/status/kew6mx25v90xylxc?svg=true)](https://ci.appveyor.com/project/miyabis/seleniumtestassist)
[![NuGet](https://img.shields.io/nuget/v/MiYABiS.SeleniumTestAssist.svg)](https://www.nuget.org/packages/MiYABiS.SeleniumTestAssist/)

It is a wrapper library to use Selenium in MSTest.  
Preparing such as Selenium is a little easier.  
The screenshot will be output to the result of MSTest when you use the method that was prepared.


How to get
==========

URL:https://www.nuget.org/packages/MiYABiS.SeleniumTestAssist/
```
PM> Install-Package MiYABiS.SeleniumTestAssist
```

Class creation template
=======================
Please add the extension of the following If you use a template.

* VS2012 Or later : [Moca.NET Templates](https://visualstudiogallery.msdn.microsoft.com/7735e52f-74f2-4ac7-8172-11cde77e6290)
* VS2010 : [Moca.NET Templates 2010](https://visualstudiogallery.msdn.microsoft.com/f97a7486-560b-425a-aa05-528dd397f5ba)


Test in IE
=======

When testing in IE requires drivers.  
Please search for , such as " iedriver " in NuGet.


Test in Chrome
=======

The same is true when you want to test in Chrome.  
Please search for , such as " chromedriver " in NuGet.

Programming
=======

Please be inherited by the test class "AbstractSeleniumTest" class.  
In the method of initialization and termination of the attributes of the test , you run the following methods.

It is already mounted if you use the " Selenium test class " of the template.

* ClassInitialize : SeleniumInitialize
* ClassCleanup : SeleniumCleanup
* TestInitialize : Selenium Browser Initialize methods
* TestCleanup : base.TestCleanup

Preparation method of Selenium is as follows .  
It runs at initialization method or in each test gave a TestInitialize attribute .

**Initialization of when running in local**

* IEInitialize
* EdgeInitialize
* FirefoxInitialize
* ChromeInitialize

**Initialization of when running in SeleniumRC**

* IERemoteInitialize
* EdgeRemoteInitialize
* FirefoxRemoteInitialize
* ChromeRemoteInitialize

[Sample](https://github.com/miyabis/SeleniumTestAssist/blob/master/WebAppSeleniumTest/UnitTest1.vb)

Screenshot
=======
The screenshot will be output to the result of MSTest when you use the method that was prepared .

```vb
Me.getScreenshot("add filename suffix")
```

[Sample](https://github.com/miyabis/SeleniumTestAssist/blob/master/WebAppSeleniumTest/UnitTest1.vb#L96)

Launch the IISExpress
=======
Please use the " IISExpressManager " class to also start IISExpress in the test.

[Sample](https://github.com/miyabis/SeleniumTestAssist/blob/master/WebAppSeleniumTest/UnitTest1.vb#L28)

License
=======

Microsoft Public License (MS-PL)

http://opensource.org/licenses/MS-PL
