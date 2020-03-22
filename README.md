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



# Mobilize.Net Collection of Information about Migrating from VB6

## [Dealing with UnResolved References](https://www.mobilize.net/blog/vbuc-expert-mode)

Whenever you're trying to convert a VB6 or Classic ASP application to .NET with the Visual Basic Upgrade Companion it is recommended that you do so in an environment where the source application can be built and executed. This will ensure all the required references are correctly registered in that environment. Sometimes, however, VBUC will still show some errors in the "Resolve References" .... 

## Moving from VB6 to WinForms

[An overview of the migrated code after upgrading from VB6 to Windows Forms](https://www.mobilize.net/blog/where-is-my-business-logic-when-migrating-to-winforms)

## Migrating VB6 applications that integrated with Office

This article provides a lot of details on how to handle project references. In particular it provides some guidance for a VB6 application that was using the EXCEL APIs  https://www.mobilize.net/blog/vb6-to-.net-missing-a-reference

## Apartment Threading

In VB6 public variables defined in standard (.bas) modules aren’t really “global variables”.
These public variables are scoped at the apartment level.
https://www.mobilize.net/blog/apartment-threading-those-little-details-that-make-a-world-of-difference-between-vb6-and-.net

## VBControlExtender

In VB6, you need VBControlExtender object for dynamically adding a control to the Controls collection using the Add method.
https://www.mobilize.net/blog/dynamically-adding-activex-in-c

## FixedLengthString and Windows API

A post with some information about how the VBUC handles fixed length strings and windows apis
https://www.mobilize.net/blog/vbuc-6.3-syntactic-optimizations

## Crystal Reports

Information about migrating Crystal Reports
 http://blogs.artinsoft.net/juan_fernando/archive/2011/09/13/Dealing-with-Crystal-Reports.aspx
 
 Crystal Reports in Windows Azure
 http://blogs.artinsoft.net/Mrojas/archive/2012/10/10/Crystal-Reports-in-Windows-Azure.aspx
 
 Using Crystal Reports in VS
 http://blogs.artinsoft.net/Mrojas/archive/2012/10/10/Use-Crystal-Reports-in-VS2010.aspx
 
## Migrating VB6 OLE Container

Quick replacement for the OLE Container Control that you had in VB6
http://blogs.artinsoft.net/Mrojas/archive/2012/01/23/Quick-replacement-for-the-VB6-OLE-Container-Control-in-NET.aspx

 
## VB6 Migration of Property Pages

http://blogs.artinsoft.net/Mrojas/archive/2009/06/09/VB6-Migration-of-Property-Pages.aspx

http://blogs.artinsoft.net/Mrojas/archive/2011/09/13/Property-Pages-in-VBNET-and-C.aspx

## How to handle Internationalization once you move to .NET

http://blogs.artinsoft.net/Mrojas/archive/2006/12/27/Taking-an-application-to-the-whole-world-(Series-1-of-3).aspx
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
