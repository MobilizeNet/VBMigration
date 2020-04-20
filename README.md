# VBMigration

![VBUC](https://www.mobilize.net/hs-fs/hubfs/1webSITE-Images/Images/Mobilize-VBUC.png)


# VBUC transforms desktop apps to .NET and web

## Preserve business logic and algorithms

Unlike a rewrite, VBUC moves existing back-end logic to the new platform, keeping proven and debugged logic and processes intact and dramatically reducing the total defects to be resolved following the migration.

## Why should I upgrade my VB 6.0 applications to .NET?

There are several drivers for a VB to .NET migration:

Integrate Windows, Web, Office and Mobile solutions
Boost system performance
Ease deployment
Improve the maintenance of an application
Increase developer productivity
Consolidate your company's valuable software assets
Avoid obsolescence of outdated software [support for VB 6 ended on April 8th, 2008 ](https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6-support-policy?redirectedfrom=MSDN).
Maintain competitive advantage

# What to see some samples ?

We are sure that you will like to see what the VBUC can do for you.

* You can take a look at our [Reference Sample Application called Salmon King Seafood (SKS)](https://github.com/MobilizeNet/SKS). This is a fictions company that sells seafood. We provide the VB6 code as well as the migrated code to C#, to .NET Core and to ASP.NET Core.
* We also prepared with Microsoft an extended sample for their [Tailwind Azure Demos](https://azureappsdemomap.com/tailwind) this demo is a Point of Sale Application that is [migrated from VB6 to Desktop and then to Azure](https://github.com/microsoft/TailwindTraders-PointOfSale).  


# Some more interestings links with Information about Migrating from VB6 to .NET or .NET Core.

## Older versions of VB

http://blogs.artinsoft.net/Mrojas/archive/2010/12/21/Upgrading-Applications-Written-in-Earlier-versions-of-Visual-Basic-for-example-VB4-VB5.aspx

## ASP Migration

The VBUC can also help you to modernize your Classic ASP WebSite.

* [Issues with COM+ when upgrading ASP to COM+](http://blogs.artinsoft.net/Mrojas/archive/2011/02/18/ASP-Migration-COM+-and-security.aspx)

* [Fixing Includes](http://blogs.artinsoft.net/Mrojas/archive/2011/01/24/ASP-to-ASPNet-Migration-Fixing-Includes.aspx)

* [Split Config Files] (http://blogs.artinsoft.net/Mrojas/archive/2011/01/06/Split-Config-files-in-several-files.aspx)

## [Dealing with UnResolved References](https://www.mobilize.net/blog/vbuc-expert-mode)

Whenever you're trying to convert a VB6 or Classic ASP application to .NET with the Visual Basic Upgrade Companion it is recommended that you do so in an environment where the source application can be built and executed. This will ensure all the required references are correctly registered in that environment. Sometimes, however, VBUC will still show some errors in the "Resolve References" .... 

## Moving from VB6 to WinForms

[An overview of the migrated code after upgrading from VB6 to Windows Forms](https://www.mobilize.net/blog/where-is-my-business-logic-when-migrating-to-winforms)

## Migrating VB6 applications that integrated with Office

This article provides a lot of details on how to handle project references. In particular it provides some guidance for a VB6 application [that was using the EXCEL APIs](https://www.mobilize.net/blog/vb6-to-.net-missing-a-reference)

## Context Sensitive Help

[How to use Context Sensitive help files in Windows Forms](
http://blogs.artinsoft.net/Mrojas/archive/2011/05/10/Converting-from-VB6-or-Winforms-to-Context-Sensitive-Help-in-Silverlight.aspx)

## Apartment Threading

In VB6 public variables defined in standard (.bas) modules aren’t really “global variables”.
These public variables are scoped [at the apartment level](
https://www.mobilize.net/blog/apartment-threading-those-little-details-that-make-a-world-of-difference-between-vb6-and-.net)

## VBControlExtender

In VB6, you need VBControlExtender object for [dynamically adding a control to the Controls collection using the Add method](
https://www.mobilize.net/blog/dynamically-adding-activex-in-c)

## FixedLengthString and Windows API and COM

A post with some information about how the VBUC handles [fixed length strings and windows apis](https://www.mobilize.net/blog/vbuc-6.3-syntactic-optimizations)

Some examples on calling [DLLs when moving from VB6 to .NET](http://blogs.artinsoft.net/Mrojas/archive/2011/05/18/Interop-Structures-to-UnManaged-Dlls.aspx)

Binary Compatibility: In VB6 when you have an ActiveX Library it was very important to use
the [BinaryCompatibility setting](http://blogs.artinsoft.net/Mrojas/archive/2010/06/24/Interop-BinaryCompatibilty-for-VB6-Migrations.aspx) to make sure that your applications did not break after a change.


Exposing [C# Classes thru Interop](http://blogs.artinsoft.net/Mrojas/archive/2010/06/23/Exposing-C-Classes-thru-Interop.aspx)

One of those subtle details is that the tlbexp tool (this is the tool that generates the .tlb for a .NET assembly)
generates a [prefix for enum elements](http://blogs.artinsoft.net/Mrojas/archive/2010/05/17/Interop-Remove-prefix-from-C-Enums-for-COM.aspx).

## Crystal Reports

* Information about migrating [Crystal Reports](http://blogs.artinsoft.net/juan_fernando/archive/2011/09/13/Dealing-with-Crystal-Reports.aspx)
  
 * [Crystal Reports in Windows Azure](http://blogs.artinsoft.net/Mrojas/archive/2012/10/10/Crystal-Reports-in-Windows-Azure.aspx)

 
 * Using [Crystal Reports in VS](http://blogs.artinsoft.net/Mrojas/archive/2012/10/10/Use-Crystal-Reports-in-VS2010.aspx)

 
## Migrating VB6 OLE Container

Quick replacement for the [OLE Container Control](http://blogs.artinsoft.net/Mrojas/archive/2012/01/23/Quick-replacement-for-the-VB6-OLE-Container-Control-in-NET.aspx) that you had in VB6


 
## VB6 Migration of Property Pages

[Migrating Property Pages](http://blogs.artinsoft.net/Mrojas/archive/2009/06/09/VB6-Migration-of-Property-Pages.aspx)

[Implementing Property Pages in VB.NET or C#](http://blogs.artinsoft.net/Mrojas/archive/2011/09/13/Property-Pages-in-VBNET-and-C.aspx)

## Windows Service in VB6 and how to upgrade it

[Creating a Windows Service in VB6](http://blogs.artinsoft.net/Mrojas/archive/2011/06/02/VB6-Windows-Service.aspx)

## How to handle Internationalization once you move to .NET

A list of some things that you should consider:

* [Internationalization of Applications](http://blogs.artinsoft.net/Mrojas/archive/2011/08/23/Things-to-consider-for-Internationalization-of-a-NET-Application.aspx)
* [Internationalization Part 2](http://blogs.artinsoft.net/Mrojas/archive/2006/12/27/Taking-an-application-to-the-whole-world-(Series-1-of-3).aspx)
* [Internationalization Part 3]
http://blogs.artinsoft.net/Mrojas/archive/2006/12/27/Starting-with-the-internationalization-bla-bla-(Part-Two).aspx



## Upgraded Stubs
When a library a library has some classes, properties, methods or events that aren't already supported an Upgrade Stub will be generated.
An Upgrade Stub is a "mock" class. 

``` C#
public class MSXML2_XMLHTTP30
{
   public string getresponseText()
   {
      UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("MSXML2.XMLHTTP30.responseText");
   return "";
   }
   public void open(string bstrMethod, string bstrUrl, object varAsync, object bstrUser, object bstrPassword)
   {
      UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("MSXML2.XMLHTTP30.open");
   }
   public void send(object varBody)
   {
      UpgradeHelpers.Helpers.NotUpgradedHelper.NotifyNotUpgradedElement("MSXML2.XMLHTTP30.send");
   }
}
```

https://www.mobilize.net/blog/vbuc-upgrade-stubs

# Videos

* A video walkthrough of our SKS reference application [migrated all the way from VB6 to the WEB](https://www.mobilize.net/webmap-demo). Excellent video by John Brown.

* [A live session using the VBUC in case you wanted a guide](https://mobilize.wistia.com/medias/xx0rrxlu98?hss_channel=fbp-200153053340106)

* [Another live session migrating Tailwind Point of Sale App to Azure]( 
https://www.youtube.com/watch?v=s8OHXA44tTg&feature=youtu.be&t=310&utm_content=105661300&utm_medium=social&utm_source=facebook&hss_channel=fbp-200153053340106)

* [10 things to know about the VBUC] (https://www.youtube.com/watch?v=XuGqQBqzePY)

* [Comparison of original VB6 app with migrated WebMAP](https://www.youtube.com/watch?v=NtfLWZhwfUk) version showing how they are visually and functionally equivalent.


* [MSDN Channel 9 Migrating VB6 apps to C# or .NET](https://channel9.msdn.com/Blogs/Bytes+by+MSDN/Bytes-by-MSDN-Dr-Ivan-Sanabria-and-Tim-Huckaby-Migrating-VB6-apps-to-C-or-NET)


* [MSDN Channel 9 Resurrecting legacy VB6 applications](https://channel9.msdn.com/Blogs/Bytes+by+MSDN/Bytes-by-MSDN-Roberto-Leiton-and-Tim-Huckaby-on-resurrecting-legacy-applications)


* Moving legacy Windows desktop applications to a modern Web architecture involves a lot of issues. [We discuss and give you ideas about how to handle them](https://www.youtube.com/watch?v=CD-uVgwiBr8).


 ## VBUC Clients Testimonials
 
 * [CFM Materials](https://www.youtube.com/watch?v=yJ398oLGBRI)

 
 * [AgWorks]( https://www.youtube.com/watch?v=xKq1ZHhOj3w)

 
 * [Hutchins]( https://www.youtube.com/watch?v=XSi9sMY1ud8&t=38s)

 
 
