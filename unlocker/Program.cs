﻿using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System;

namespace unlocker
{
    using Properties;

    public static class Program
    {
        public static readonly string[] OSSPP_PATHS =
        {
            @"%ProgramFiles%\Microsoft Office\Office15\ospp.vbs",
            @"%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs",
            @"%ProgramFiles%\Microsoft Office\Office16\ospp.vbs",
            @"%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs",
        };
        public static readonly string[] KMS_SERVERS =
        {
            "kms7.MSGuides.com",
            "kms8.MSGuides.com",
            "kms9.MSGuides.com",
        };
        public static readonly string KEY_SERIAL = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99";
        public static readonly string[] KEY_PARTS = {
            "WFG99",
            "DRTFM",
            "BTDRB",
            "CPQVG"
        };
        public static readonly Dictionary<string, byte[]> LICENSES = new Dictionary<string, byte[]>
        {
            ["ProPlusVL_KMS_Client-ppd.xrm-ms"] = Resources.ProPlusVL_KMS_Client_ppd,
            ["ProPlusVL_KMS_Client-ul.xrm-ms"] = Resources.ProPlusVL_KMS_Client_ul,
            ["ProPlusVL_MAK-ppd.xrm-ms"] = Resources.ProPlusVL_MAK_ppd,
            ["ProPlusVL_MAK-ul-phn.xrm-ms"] = Resources.ProPlusVL_MAK_ul_phn,
            ["ProPlusVL_KMS_Client-ul-oob.xrm-ms"] = Resources.ProPlusVL_KMS_Client_ul_oob,
            ["ProPlusVL_MAK-pl.xrm-ms"] = Resources.ProPlusVL_MAK_pl,
            ["ProPlusVL_MAK-ul-oob.xrm-ms"] = Resources.ProPlusVL_MAK_ul_oob,
        };

        public static readonly Regex REGEX_PROPLUSVL = new Regex(@"^proplusvl_(kms|mak).*\.xrm-ms$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        public static readonly Regex REGEX_LICDIR = new Regex(@"Licenses\d+", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static void Main()
        {
            try
            {
                Console.BufferWidth = Math.Max(Console.BufferWidth, 72);
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine(@"
======================================================================
              _                               __    __  ____   __   
  _   _ _ __ | | ___ __   _____      ___ __  / /_  / /_| ___| / /_  
 | | | | '_ \| |/ / '_ \ / _ \ \ /\ / / '_ \| '_ \| '_ \___ \| '_ \ 
 | |_| | | | |   <| | | | (_) \ V  V /| | | | (_) | (_) |__) | (_) |
  \__,_|_| |_|_|\_\_| |_|\___/ \_/\_/ |_| |_|\___/ \___/____/ \___/ 

  COPYRIGHT (C) 2020, unknown6656 -- https://github.com/unknown6656
======================================================================
");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Supported products:");
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine(@"
- Microsoft Office Standard 2016
- Microsoft Office Standard 2019
- Microsoft Office Professional Plus 2016
- Microsoft Office Professional Plus 2019
".Trim());

                FileInfo ospp = null;
                bool activated = false;

                foreach (string path in OSSPP_PATHS)
                    try
                    {
                        FileInfo fi = new FileInfo(Environment.ExpandEnvironmentVariables(path));

                        if (fi?.Exists ?? false)
                        {
                            ospp = fi;

                            break;
                        }
                    }
                    catch
                    {
                    }

                if (ospp is null)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("The Microsoft Office installation could not be found. Please verify and/or repair your installation.");

                    return;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("KMS Script location: " + ospp.FullName);
                }

                Tuple<int, string> exec_ospp(string args)
                {
                    using (Process proc = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            WorkingDirectory = ospp.Directory.FullName,
                            FileName = Environment.ExpandEnvironmentVariables(@"%windir%\system32\cscript.exe"),
                            Arguments = $"/nologo \"{ospp.FullName}\" {args}".Trim(),
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                        }
                    })
                    {
                        proc.Start();
                        proc.WaitForExit();

                        string stdout = proc.StandardOutput.ReadToEnd().Replace("\r\n", "\n").Trim() + "\n";
                        string stderr = proc.StandardError.ReadToEnd().Trim() + "\n";

                        stdout = Regex.Replace(stdout, @"^---[^\n\r]*---$", "", RegexOptions.Multiline).Replace("\n\n", "\n").Trim();

                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine($"Excuting '%windir%\\system32\\cscript {proc.StartInfo.Arguments}' ...");

                        if (!string.IsNullOrWhiteSpace(stdout))
                        {
                            Console.ForegroundColor = ConsoleColor.DarkCyan;
                            Console.WriteLine(stdout);
                        }

                        if (!string.IsNullOrWhiteSpace(stderr))
                        {
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine(stderr);
                        }

                        return new Tuple<int, string>(proc.ExitCode, stdout);
                    }
                }
                DirectoryInfo licensedir = new DirectoryInfo($"{ospp.Directory.FullName}/../root");

                if (!licensedir.Exists)
                    licensedir.Create();

                licensedir = licensedir.EnumerateDirectories().FirstOrDefault(dir => REGEX_LICDIR.IsMatch(dir.Name)) ?? licensedir.CreateSubdirectory("Licenses16");

                foreach (string licence in LICENSES.Keys)
                {
                    string licpath = Path.Combine(licensedir.FullName, licence);

                    if (!File.Exists(licpath))
                    {
                        Console.ForegroundColor = ConsoleColor.Gray;
                        Console.WriteLine($"Generating license file '{licpath}' from internal storage ...");

                        File.WriteAllBytes(licpath, LICENSES[licence]);

                        Console.WriteLine($"Written license file '{licpath}' from internal storage ({LICENSES[licence].Length} Bytes).");
                    }
                }

                foreach (FileInfo licence in licensedir.GetFiles())
                    if (REGEX_PROPLUSVL.IsMatch(licence.Name))
                        exec_ospp($"/inslic:\"{licence.FullName}\"");

                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Activating Microsoft Office ...");

                foreach (string part in KEY_PARTS)
                    exec_ospp($"/unpkey:{part}");

                exec_ospp($"/inpkey:{KEY_SERIAL}");

                foreach (string server in KMS_SERVERS)
                {
                    exec_ospp($"/sethst:{server}");

                    Tuple<int, string> result = exec_ospp("/act");

                    if (result.Item2.ToLower().Contains("successful"))
                    {
                        activated = true;

                        break;
                    }
                }

                Console.WriteLine();

                if (activated)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("    MICROSOFT OFFICE SUCCESSFULLY ACTIVATED!");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("    MICROSOFT OFFICE COULD NOT BE ACTIVATED.");
                }
            }
            finally
            {
                Console.ForegroundColor = ConsoleColor.Gray;
            }
        }
    }
}

/*
 * BATCH EQUIVALENT:

    @echo off
    (if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")
    (if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")
    (for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > nul)
    (for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > nul)
    cscript //nologo ospp.vbs /unpkey:WFG99 > nul
    cscript //nologo ospp.vbs /unpkey:DRTFM > nul
    cscript //nologo ospp.vbs /unpkey:BTDRB > nul
    cscript //nologo ospp.vbs /unpkey:CPQVG > nul
    cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 > nul
    set i=1
:server
    if %i%==1 set KMS_Sev=kms7.MSGuides.com
    if %i%==2 set KMS_Sev=kms8.MSGuides.com
    if %i%==3 set KMS_Sev=kms9.MSGuides.com
    if %i%==4 goto halt
    cscript //nologo ospp.vbs /sethst:%KMS_Sev% > nul
    cscript //nologo ospp.vbs /act | find /i "successful" || (set /a i+=1 & goto server)
    goto halt
:halt
    echo "an error occured."
*/
