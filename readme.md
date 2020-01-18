# Microsoft Office 2016 Unlocker

This repository contains the code and the compiled binary to active Microsoft Office 2016. The binary `unlocker.exe` depends on *any* one of the following dependencies to run:
 - .NET-Core 2.0 or higher
 - .NET-Framework 4.6.1 or higher
 - Mono 5.4 or higher
 - UWP 10.0.16299 or higher

Usually, Windows comes with one of these dependencies pre-installed, however, one can find the corresponding downloads at https://dotnet.microsoft.com/download.

--------------

Virustotal reports the file to be **clean**:
 - URL scan: https://www.virustotal.com/gui/url/069296951d192e81ef3914016602801d919c905536117f990e33ec251e4d6b7d/detection
 - File scan: https://www.virustotal.com/gui/file/9192fc4263ed1917dc791c65483d1342206fa8186dcaf5d687235907f91054b8/detection

However, Windows Defender might be raising the alarm on the **source file** `Program.cs`, instead of the compiled binary. So, add your code directory to the Windows Defender exclusion list if you decide to build the application from source.
