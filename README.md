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

|:-|:-|
|ClassInitialize Attribute|SeleniumInitialize()|
|ClassCleanup Attribute|SeleniumCleanup()|
|TestInitialize Attribute|Selenium Browser Initialize method|
|TestCleanup Attribute|base.TestCleanup()|


License
=======

Microsoft Public License (MS-PL)

http://opensource.org/licenses/MS-PL
