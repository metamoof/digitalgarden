---
{"dg-publish":true,"permalink":"/blog/invoking-new-outlook/"}
---

I was trying to work out how to add a button to open the Calendar in New Outlook (the replacement for Outlook, which basically seems to be a shell round the outlook.office.com PWA) to my Stream Deck. As such I needed a command line I could run from PowerShell. This turns out to be more complicated than it seems.

It's definitely possible, thought, there's a Jump Menu option to do it:

![Pasted image 20231107075856.png](/img/user/Blog/attachments/Pasted%20image%2020231107075856.png)

## New Outlook is not a normal windows executable

In the olden days, you'd just invoke outlook.exe with [command line switches](https://support.microsoft.com/en-gb/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6#Category=Outlook). However, when I try this, it gets me to the Old Outlook, and I want New Outlook. 

It turns out that New Outlook is actually a Universal Windows Platform (UWP) app of the type distributed with the Microsoft Store.

## Finding out how to invoke New Outlook - the UWP Shortcut

By going to [shell:AppsFolder](shell:AppsFolder) in explorer, you can see the UWP applications list, scroll to Outlook (New), right click and Create a Shortcut. This will create .lnk file in the desktop:

![Pasted image 20231107065643.png](/img/user/Zettelkasen/attachments/Pasted%20image%2020231107065643.png)

double-clicking or invoking this link will open outlook.

## Listing UWP apps in powershell

To get this on powershell:

```powershell
❯ Get-AppxPackage | Where-Object Name -Like "*utlook*"

RunspaceId             : 999a8d9a-04d0-4cee-b3ff-9a2072661058
Name                   : Microsoft.OutlookForWindows
Publisher              : CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US
PublisherId            : 8wekyb3d8bbwe
Architecture           : X64
ResourceId             :
Version                : 1.2023.1018.300
PackageFamilyName      : Microsoft.OutlookForWindows_8wekyb3d8bbwe
PackageFullName        : Microsoft.OutlookForWindows_1.2023.1018.300_x64__8wekyb3d8bbwe
InstallLocation        : C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2023.1018.300_x64__8wekyb3d8bbwe
IsFramework            : False
PackageUserInformation : {}
IsResourcePackage      : False
IsBundle               : False
IsDevelopmentMode      : False
NonRemovable           : False
Dependencies           : {}
IsPartiallyStaged      : False
SignatureKind          : Developer
Status                 : Ok
```

The PackageFamilyName is the same as that shortcut link

However, as per [How to Open a Windows Modern UWP App From the Command Line | The Poet Engineer (postach.io)](https://poetengineer.postach.io/post/how-to-open-windows-modern-app-from-the-command-line) we also need to find the App Name in order to create the URL. We can do this with the PackageFullName from above

```powershell
❯ (Get-AppxPackageManifest Microsoft.OutlookForWindows_1.2023.1018.300_x64__8wekyb3d8bbwe).package.Applications.Application

EntryPoint     : Windows.FullTrustApplication
Executable     : olk.exe
Id             : Microsoft.OutlookforWindows
VisualElements : VisualElements
Extensions     : Extensions
```

The Appname here is thus Microsoft.OutlookforWindows.

We can now compose the full command into [shell:AppsFolder\Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookForWindows](shell:AppsFolder\Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookForWindows)and run it using:

```powershell
> Start-Process shell:AppsFolder\Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookForWindows
```

This, however won't show the calendar.

## Getting to the calendar

From the Application above we see the executable is called `olk.exe`- we can also discover this from the task manager.

As per [reddit](https://www.reddit.com/r/PowerShell/comments/fyh5lh/can_powershell_launch_a_jump_list_item/), if we open the calendar from the Taskbar Jump Menu, we can then  query the Windows CMI to get the command line:

```powershell
❯ Get-CimInstance Win32_process | Where-Object {$_.Name -eq 'olk.exe'} | Select-Object CommandLine | Out-String -width 250

CommandLine
-----------
"C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.2023.1018.300_x64__8wekyb3d8bbwe\olk.exe" ms-outlook:launch?calendar
```
That's how you can invoke it:

```powershell
❯ Start-Process shell:Appsfolder\Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookforWindows ms-outlook:launch?calendar
```

## It's a URI!

But wait, that last bit looks suspiciously like a URI.

In fact, it is, and just running [ms-outlook:launch?calendar](ms-outlook:launch?calendar) will do the same thing.

I can't seem to find much information on the `ms-outlook` URI scheme, [it doesn't appear to be documented officially](https://learn.microsoft.com/en-us/office/client-developer/office-uri-schemes), but it does look like there are some fun things that can be done with it. I'll have to experiment with whether it's similar to the PWA URLs