# VBMigration
Mobilize.Net Collection of Information about Migrating from VB6

[Dealing with UnResolved References](https://www.mobilize.net/blog/vbuc-expert-mode)

Whenever you're trying to convert a VB6 or Classic ASP application to .NET with the Visual Basic Upgrade Companion it is recommended that you do so in an environment where the source application can be built and executed. This will ensure all the required references are correctly registered in that environment. Sometimes, however, VBUC will still show some errors in the "Resolve References" .... 

Moving from VB6 to WinForms

[An overview of the migrated code after upgrading from VB6 to Windows Forms](https://www.mobilize.net/blog/where-is-my-business-logic-when-migrating-to-winforms)

Migrating VB6 applications that integrated with Office

This article provides a lot of details on how to handle project references. In particular it provides some guidance for a VB6 application that was using the EXCEL APIs  https://www.mobilize.net/blog/vb6-to-.net-missing-a-reference

Upgraded Stubs
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
