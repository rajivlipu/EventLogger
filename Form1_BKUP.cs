﻿using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using System.Dynamic;
using System.IO.Packaging;
using System.Runtime.InteropServices.ComTypes;

namespace EventLogger
{
    public partial class Form1 : Form
    {

        #region Set Threshold
        //Set Threshold in Seconds
        const int SetTimerTime = 15;
        const int SetReChkEvtTime = 3600;

        const int SetCPUTime = 30;
        const int SetMemTime = 30;
        const int SetDiskTime = 30;

        const decimal setCpu = 91;
        const decimal setMem = 90;
        const decimal setDisk = 90;
        const decimal setHDD = 5;

        //Set Thread Threshold in MS
        const int SetResTrackingTime = 5000;
        const int SetEVTTrackingTime = 60000;
        const int SetSNTrackingTime = 30000;

        #endregion




        DispatcherTimer Timer99 = new DispatcherTimer();
        static PerformanceCounter myCounter = new PerformanceCounter("PhysicalDisk", "% Disk Time", "_Total");
        static PerformanceCounter myCounter1 = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        static List<data> MemData = new List<data>();
        static List<data> CpuData = new List<data>();
        static List<data> DiskData = new List<data>();

        static Dictionary<string ,DateTime> EvtTrack= new Dictionary<string,DateTime>();



        static string aStr = "";
        static List<Dictionary<int, DateTime>> keyValuePairs = new List<Dictionary<int, DateTime>>();

        public Form1()
        {
            InitializeComponent();
            Thread ThreadObject1 = new Thread(GetEvt);
            ThreadObject1.Start();
            Thread ThreadObject2 = new Thread(GetRes);
            ThreadObject2.Start();
            Thread ThreadObject3 = new Thread(GetSNT);
            ThreadObject3.Start();

            Timer99.Tick += Timer99_Tick;
            Timer99.Interval = new TimeSpan(0, 0, 0, SetTimerTime, 0);
            Timer99.IsEnabled = true;

            
        }

        #region formevents
        /// <summary>
        /// Set event time to restrict the duplicate LLM call
        /// </summary>
        /// <param name="Evt"></param>
        /// <returns></returns>
        public void SetEvtTime(string Evt)
        {
            DateTime EvtTime = DateTime.Now;
            if (EvtTrack.ContainsKey(Evt))
            {
                EvtTime = EvtTrack[Evt];
                TimeSpan span = EvtTime.Subtract(DateTime.Now);
                if (span.TotalSeconds > SetReChkEvtTime)
                {
                    EvtTrack.Remove(Evt);
                    EvtTrack.Add(Evt, DateTime.Now);
                }
            }
            else
            {
                EvtTrack.Add(Evt, DateTime.Now);
            }
        }
        /// <summary>
        /// Get Event Time
        /// </summary>
        /// <param name="Evt"></param>
        /// <returns></returns>
        public bool GetEvtTime(string Evt)
        {
            bool TrigLLM = false;
            DateTime EvtTime = DateTime.Now;
            if (EvtTrack.ContainsKey(Evt))
            {
                EvtTime = EvtTrack[Evt];
                TimeSpan span = EvtTime.Subtract(DateTime.Now);
                if (span.TotalSeconds > SetReChkEvtTime)
                {
                    TrigLLM = true;
                }
                else
                {
                    TrigLLM = false;
                }
            }
            else
            {
                TrigLLM = true;
            }
            return TrigLLM;
        }
        /// <summary>
        /// Timer event to track the system events to trigger LLM call
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Timer99_Tick(object sender, EventArgs e)
        {
            var drive = new DriveInfo("c");
            long freeSpaceInBytes = drive.AvailableFreeSpace;
            long TotalSpaceInBytes = drive.TotalSize;
            long percentage = (freeSpaceInBytes * 100) / TotalSpaceInBytes;

            if (percentage < setHDD)
            {
                if (GetEvtTime("HDD"))
                {
                    textBox1.Text = "Available Disk space in C: " + percentage.ToString();
                    long harddiskper = 100 - percentage;
                    SetEvtTime("HDD");
                    await callLLMAsync("Hard disk utilization is " + harddiskper.ToString() + " percentage", true);
                }
            }
            if (CpuData.Count > 0)
            {
                if (GetEvtTime("CPU"))
                {
                    textBox1.Text = "CPU" + CpuData[0].doc.ToShortTimeString() + ":" + CpuData[0].value.ToString() + Environment.NewLine + CpuData[CpuData.Count - 1].doc.ToShortTimeString() + ":" + CpuData[CpuData.Count - 1].value.ToString() + Environment.NewLine;
                    TimeSpan span = CpuData[CpuData.Count - 1].doc.Subtract(CpuData[0].doc);
                    if (span.TotalSeconds > SetCPUTime)
                    {
                        SetEvtTime("CPU");
                        await callLLMAsync("CPU usage greater than 90 percentage for more than 30 second", true);
                    }
                }
            }
            if (MemData.Count > 0)
            {
                if (GetEvtTime("MEM"))
                {
                    textBox1.Text += "Mem" + MemData[0].doc.ToShortTimeString() + ":" + MemData[0].value.ToString() + Environment.NewLine + MemData[MemData.Count - 1].doc.ToShortTimeString() + ":" + MemData[MemData.Count - 1].value.ToString() + Environment.NewLine;
                    TimeSpan span = MemData[MemData.Count - 1].doc.Subtract(MemData[0].doc);
                    if (span.TotalSeconds > SetMemTime)
                    {
                        SetEvtTime("MEM");
                        await callLLMAsync("Memory usage greater than 90 percentage for more than 30 second", true);
                    }
                }
            }
            if (DiskData.Count > 0)
            {
                if (GetEvtTime("DISK"))
                {
                    textBox1.Text += "Disk" + DiskData[0].doc.ToShortTimeString() + ":" + DiskData[0].value.ToString() + Environment.NewLine + DiskData[DiskData.Count - 1].doc.ToShortTimeString() + ":" + DiskData[DiskData.Count - 1].value.ToString() + Environment.NewLine;
                    TimeSpan span = DiskData[DiskData.Count - 1].doc.Subtract(DiskData[0].doc);
                    if (span.TotalSeconds > SetDiskTime)
                    {
                        SetEvtTime("DISK");
                        await callLLMAsync("Disk usage greater than 90 percentage for more than 30 second", true);
                    }
                }
            }
            textBox1.Text = aStr + textBox1.Text;
        }
        /// <summary>
        /// Form Load 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
        }
        /// <summary>
        /// Code to restrict closing the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }
        #endregion

        
        #region action events
        /// <summary>
        /// restart system
        /// </summary>
        public void RestartPC() {
            ExecuteCommand("Restart-Computer");
        }
        /// <summary>
        /// disk cleanup
        /// </summary>
        public void diskCleanup()
        {

            string aCmd = @"write-Host ""Disk Cleanup Started"" -ForegroundColor Red  
$Path = 'C' + ':\$Recycle.Bin'
Get - ChildItem $Path - Force - Recurse - ErrorAction SilentlyContinue |
Remove - Item - Recurse - Exclude *.ini - ErrorAction SilentlyContinue


$Path1 = 'C' + ':\Windows\Temp'
Get - ChildItem $Path1 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue
$Path2 = 'C' + ':\Windows\Prefetch'
Get - ChildItem $Path2 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue
$Path3 = 'C' + ':\Users\*\AppData\Local\Temp'
Get - ChildItem $Path4 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue

cleanmgr / sagerun:1 | out-Null
Sleep 5
write-Host ""Disk Cleanup Successfully done"" -ForegroundColor Green  
Sleep 10";
            ExecuteCommand(aCmd);
        }
        /// <summary>
        /// temp file clean
        /// </summary>
        public void tempClean()
        {

            string aCmd = @"
$Path1 = 'C' + ':\Windows\Temp'
Get - ChildItem $Path1 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue
$Path2 = 'C' + ':\Windows\Prefetch'
Get - ChildItem $Path2 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue
$Path3 = 'C' + ':\Users\*\AppData\Local\Temp'
Get - ChildItem $Path4 - Force - Recurse - ErrorAction SilentlyContinue | Remove - Item - Recurse - Force - ErrorAction SilentlyContinue";
            ExecuteCommand(aCmd);
        }
        /// <summary>
        /// cache clean
        /// </summary>
        public void cacheClean()
        {

            string aCmd = @"#------------------------------------------------------------------#
#- Clear-GlobalWindowsCache                                        #
#------------------------------------------------------------------#
Function Clear-GlobalWindowsCache {
    Remove-CacheFiles 'C:\Windows\Temp' 
    Remove-CacheFiles ""C:\`$Recycle.Bin""
    Remove-CacheFiles ""C:\Windows\Prefetch""
    C:\Windows\System32\rundll32.exe InetCpl.cpl, ClearMyTracksByProcess 255
    C:\Windows\System32\rundll32.exe InetCpl.cpl, ClearMyTracksByProcess 4351
}

#------------------------------------------------------------------#
#- Clear-UserCacheFiles                                            #
#------------------------------------------------------------------#
Function Clear-UserCacheFiles {
    # Stop-BrowserSessions
    ForEach($localUser in (Get-ChildItem 'C:\users').Name)
    {
        Clear-ChromeCache $localUser
        Clear-EdgeCacheFiles $localUser
        Clear-FirefoxCacheFiles $localUser
        Clear-WindowsUserCacheFiles $localUser
        Clear-TeamsCacheFiles $localUser
    }
}

#------------------------------------------------------------------#
#- Clear-WindowsUserCacheFiles                                     #
#------------------------------------------------------------------#
Function Clear-WindowsUserCacheFiles {
    param([string]$user=$env:USERNAME)
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Temp""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\WER""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\INetCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\INetCookies""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IECompatCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IECompatUaCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IEDownloadHistory""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\Temporary Internet Files""    
}

#Region HelperFunctions

#------------------------------------------------------------------#
#- Stop-BrowserSessions                                            #
#------------------------------------------------------------------#
Function Stop-BrowserSessions {
   $activeBrowsers = Get-Process Firefox*,Chrome*,Waterfox*,Edge*
   ForEach($browserProcess in $activeBrowsers)
   {
       try 
       {
           $browserProcess.CloseMainWindow() | Out-Null 
       } catch { }
   }
}

#------------------------------------------------------------------#
#- Get-StorageSize                                                 #
#------------------------------------------------------------------#
Function Get-StorageSize {
    Get-WmiObject Win32_LogicalDisk | 
    Where-Object { $_.DriveType -eq ""3"" } | 
    Select-Object SystemName, 
        @{ Name = ""Drive"" ; Expression = { ( $_.DeviceID ) } },
        @{ Name = ""Size (GB)"" ; Expression = {""{0:N1}"" -f ( $_.Size / 1gb)}},
        @{ Name = ""FreeSpace (GB)"" ; Expression = {""{0:N1}"" -f ( $_.Freespace / 1gb ) } },
        @{ Name = ""PercentFree"" ; Expression = {""{0:P1}"" -f ( $_.FreeSpace / $_.Size ) } } |
    Format-Table -AutoSize | Out-String
}

#------------------------------------------------------------------#
#- Remove-CacheFiles                                               #
#------------------------------------------------------------------#
Function Remove-CacheFiles {
    param([Parameter(Mandatory=$true)][string]$path)    
    BEGIN 
    {
        $originalVerbosePreference = $VerbosePreference
        $VerbosePreference = 'Continue'  
    }
    PROCESS 
    {
        if((Test-Path $path))
        {
            if([System.IO.Directory]::Exists($path))
            {
                try 
                {
                    if($path[-1] -eq '\')
                    {
                        [int]$pathSubString = $path.ToCharArray().Count - 1
                        $sanitizedPath = $path.SubString(0, $pathSubString)
                        Remove-Item -Path ""$sanitizedPath\*"" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
                    }
                    else 
                    {
                        Remove-Item -Path ""$path\*"" -Recurse -Force -ErrorAction SilentlyContinue -Verbose              
                    } 
                } catch { }
            }
            else 
            {
                try 
                {
                    Remove-Item -Path $path -Force -ErrorAction SilentlyContinue -Verbose
                } catch { }
            }
        }    
    }
    END 
    {
        $VerbosePreference = $originalVerbosePreference
    }
}

#Endregion HelperFunctions

#Region Browsers

#Region ChromiumBrowsers

#------------------------------------------------------------------#
#- Clear-ChromeCache                                               #
#------------------------------------------------------------------#
Function Clear-ChromeCache {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Google\Chrome\User Data\Default""))
    {
        $chromeAppData = ""C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default"" 
        $possibleCachePaths = @('Cache','Cache2\entries\','Cookies','History','Top Sites','VisitedLinks','Web Data','Media Cache','Cookies-Journal','ChromeDWriteFontCache')
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$chromeAppData\$cachePath""
        }      
    } 
}

#------------------------------------------------------------------#
#- Clear-EdgeCache                                                 #
#------------------------------------------------------------------#
Function Clear-EdgeCache {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\Users$user\AppData\Local\Microsoft\Edge\User Data\Default""))
    {
        $EdgeAppData = ""C:\Users$user\AppData\Local\Microsoft\Edge\User Data\Default""
        $possibleCachePaths = @('Cache','Cache2\entries','Cookies','History','Top Sites','Visited Links','Web Data','Media History','Cookies-Journal')
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$EdgeAppData$cachePath""
        }
        }
}

#Endregion ChromiumBrowsers

#Region FirefoxBrowsers

#------------------------------------------------------------------#
#- Clear-FirefoxCacheFiles                                         #
#------------------------------------------------------------------#
Function Clear-FirefoxCacheFiles {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Mozilla\Firefox\Profiles""))
    {
        $possibleCachePaths = @('cache','cache2\entries','thumbnails','cookies.sqlite','webappsstore.sqlite','chromeappstore.sqlite')
        $firefoxAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Local\Mozilla\Firefox\Profiles"" | Where-Object { $_.Name -match 'Default' }[0]).FullName 
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$firefoxAppDataPath\$cachePath""
        }
    } 
}

#------------------------------------------------------------------#
#- Clear-WaterfoxCacheFiles                                        #
#------------------------------------------------------------------#
Function Clear-WaterfoxCacheFiles { 
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Waterfox\Profiles""))
    {
        $possibleCachePaths = @('cache','cache2\entries','thumbnails','cookies.sqlite','webappsstore.sqlite','chromeappstore.sqlite')
        $waterfoxAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Local\Waterfox\Profiles"" | Where-Object { $_.Name -match 'Default' }[0]).FullName
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$waterfoxAppDataPath\$cachePath""
        }
    }   
}

#Endregion FirefoxBrowsers

#Endregion Browsers

#Region CommunicationPlatforms

#------------------------------------------------------------------#
#- Clear-TeamsCacheFiles                                           #
#------------------------------------------------------------------#
Function Clear-TeamsCacheFiles { 
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Roaming\Microsoft\Teams""))
    {
        $possibleCachePaths = @('cache','blob_storage','databases','gpucache','Indexeddb','Local Storage','application cache\cache')
        $teamsAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Roaming\Microsoft\Teams"" | Where-Object { $_.Name -match 'Default' }[0]).FullName
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$teamsAppDataPath\$cachePath""
        }
    }   
}

#Endregion CommunicationPlatforms

#------------------------------------------------------------------#
#- MAIN                                                            #
#------------------------------------------------------------------#

$StartTime = (Get-Date)

Get-StorageSize

Clear-UserCacheFiles
Clear-GlobalWindowsCache

Get-StorageSize

$EndTime = (Get-Date)
Write-Verbose ""Elapsed Time: $(($StartTime - $EndTime).totalseconds) seconds""";
            ExecuteCommand(aCmd);
        }

        /// <summary>
        /// browser data clean
        /// </summary>
    public void browserClean()
        {

            string aCmd = @"#------------------------------------------------------------------#
#- Clear-GlobalWindowsCache                                        #
#------------------------------------------------------------------#
Function Clear-GlobalWindowsCache {
    Remove-CacheFiles 'C:\Windows\Temp' 
    Remove-CacheFiles ""C:\`$Recycle.Bin""
    Remove-CacheFiles ""C:\Windows\Prefetch""
    C:\Windows\System32\rundll32.exe InetCpl.cpl, ClearMyTracksByProcess 255
    C:\Windows\System32\rundll32.exe InetCpl.cpl, ClearMyTracksByProcess 4351
}

#------------------------------------------------------------------#
#- Clear-UserCacheFiles                                            #
#------------------------------------------------------------------#
Function Clear-UserCacheFiles {
    # Stop-BrowserSessions
    ForEach($localUser in (Get-ChildItem 'C:\users').Name)
    {
        Clear-ChromeCache $localUser
        Clear-EdgeCacheFiles $localUser
        Clear-FirefoxCacheFiles $localUser
        #Clear-WindowsUserCacheFiles $localUser
        #Clear-TeamsCacheFiles $localUser
    }
}

#------------------------------------------------------------------#
#- Clear-WindowsUserCacheFiles                                     #
#------------------------------------------------------------------#
Function Clear-WindowsUserCacheFiles {
    param([string]$user=$env:USERNAME)
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Temp""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\WER""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\INetCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\INetCookies""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IECompatCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IECompatUaCache""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\IEDownloadHistory""
    Remove-CacheFiles ""C:\Users\$user\AppData\Local\Microsoft\Windows\Temporary Internet Files""    
}

#Region HelperFunctions

#------------------------------------------------------------------#
#- Stop-BrowserSessions                                            #
#------------------------------------------------------------------#
Function Stop-BrowserSessions {
   $activeBrowsers = Get-Process Firefox*,Chrome*,Waterfox*,Edge*
   ForEach($browserProcess in $activeBrowsers)
   {
       try 
       {
           $browserProcess.CloseMainWindow() | Out-Null 
       } catch { }
   }
}

#------------------------------------------------------------------#
#- Get-StorageSize                                                 #
#------------------------------------------------------------------#
Function Get-StorageSize {
    Get-WmiObject Win32_LogicalDisk | 
    Where-Object { $_.DriveType -eq ""3"" } | 
    Select-Object SystemName, 
        @{ Name = ""Drive"" ; Expression = { ( $_.DeviceID ) } },
        @{ Name = ""Size (GB)"" ; Expression = {""{0:N1}"" -f ( $_.Size / 1gb)}},
        @{ Name = ""FreeSpace (GB)"" ; Expression = {""{0:N1}"" -f ( $_.Freespace / 1gb ) } },
        @{ Name = ""PercentFree"" ; Expression = {""{0:P1}"" -f ( $_.FreeSpace / $_.Size ) } } |
    Format-Table -AutoSize | Out-String
}

#------------------------------------------------------------------#
#- Remove-CacheFiles                                               #
#------------------------------------------------------------------#
Function Remove-CacheFiles {
    param([Parameter(Mandatory=$true)][string]$path)    
    BEGIN 
    {
        $originalVerbosePreference = $VerbosePreference
        $VerbosePreference = 'Continue'  
    }
    PROCESS 
    {
        if((Test-Path $path))
        {
            if([System.IO.Directory]::Exists($path))
            {
                try 
                {
                    if($path[-1] -eq '\')
                    {
                        [int]$pathSubString = $path.ToCharArray().Count - 1
                        $sanitizedPath = $path.SubString(0, $pathSubString)
                        Remove-Item -Path ""$sanitizedPath\*"" -Recurse -Force -ErrorAction SilentlyContinue -Verbose
                    }
                    else 
                    {
                        Remove-Item -Path ""$path\*"" -Recurse -Force -ErrorAction SilentlyContinue -Verbose              
                    } 
                } catch { }
            }
            else 
            {
                try 
                {
                    Remove-Item -Path $path -Force -ErrorAction SilentlyContinue -Verbose
                } catch { }
            }
        }    
    }
    END 
    {
        $VerbosePreference = $originalVerbosePreference
    }
}

#Endregion HelperFunctions

#Region Browsers

#Region ChromiumBrowsers

#------------------------------------------------------------------#
#- Clear-ChromeCache                                               #
#------------------------------------------------------------------#
Function Clear-ChromeCache {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Google\Chrome\User Data\Default""))
    {
        $chromeAppData = ""C:\Users\$user\AppData\Local\Google\Chrome\User Data\Default"" 
        $possibleCachePaths = @('Cache','Cache2\entries\','Cookies','History','Top Sites','VisitedLinks','Web Data','Media Cache','Cookies-Journal','ChromeDWriteFontCache')
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$chromeAppData\$cachePath""
        }      
    } 
}

#------------------------------------------------------------------#
#- Clear-EdgeCache                                                 #
#------------------------------------------------------------------#
Function Clear-EdgeCache {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\Users$user\AppData\Local\Microsoft\Edge\User Data\Default""))
    {
        $EdgeAppData = ""C:\Users$user\AppData\Local\Microsoft\Edge\User Data\Default""
        $possibleCachePaths = @('Cache','Cache2\entries','Cookies','History','Top Sites','Visited Links','Web Data','Media History','Cookies-Journal')
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$EdgeAppData$cachePath""
        }
        }
}

#Endregion ChromiumBrowsers

#Region FirefoxBrowsers

#------------------------------------------------------------------#
#- Clear-FirefoxCacheFiles                                         #
#------------------------------------------------------------------#
Function Clear-FirefoxCacheFiles {
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Mozilla\Firefox\Profiles""))
    {
        $possibleCachePaths = @('cache','cache2\entries','thumbnails','cookies.sqlite','webappsstore.sqlite','chromeappstore.sqlite')
        $firefoxAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Local\Mozilla\Firefox\Profiles"" | Where-Object { $_.Name -match 'Default' }[0]).FullName 
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$firefoxAppDataPath\$cachePath""
        }
    } 
}

#------------------------------------------------------------------#
#- Clear-WaterfoxCacheFiles                                        #
#------------------------------------------------------------------#
Function Clear-WaterfoxCacheFiles { 
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Local\Waterfox\Profiles""))
    {
        $possibleCachePaths = @('cache','cache2\entries','thumbnails','cookies.sqlite','webappsstore.sqlite','chromeappstore.sqlite')
        $waterfoxAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Local\Waterfox\Profiles"" | Where-Object { $_.Name -match 'Default' }[0]).FullName
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$waterfoxAppDataPath\$cachePath""
        }
    }   
}

#Endregion FirefoxBrowsers

#Endregion Browsers

#Region CommunicationPlatforms

#------------------------------------------------------------------#
#- Clear-TeamsCacheFiles                                           #
#------------------------------------------------------------------#
Function Clear-TeamsCacheFiles { 
    param([string]$user=$env:USERNAME)
    if((Test-Path ""C:\users\$user\AppData\Roaming\Microsoft\Teams""))
    {
        $possibleCachePaths = @('cache','blob_storage','databases','gpucache','Indexeddb','Local Storage','application cache\cache')
        $teamsAppDataPath = (Get-ChildItem ""C:\users\$user\AppData\Roaming\Microsoft\Teams"" | Where-Object { $_.Name -match 'Default' }[0]).FullName
        ForEach($cachePath in $possibleCachePaths)
        {
            Remove-CacheFiles ""$teamsAppDataPath\$cachePath""
        }
    }   
}

#Endregion CommunicationPlatforms

#- MAIN                                                            #


Clear-UserCacheFiles
";
            ExecuteCommand(aCmd);
        }

        #endregion


        #region functions

        /// <summary>
        /// Get System Resource
        /// </summary>
        static void GetRes()
        {
            Thread.Sleep(SetResTrackingTime);
            decimal diskUsage = Convert.ToDecimal(myCounter.NextValue());
            decimal a1 = Convert.ToDecimal(myCounter1.NextValue());
            if (a1 > setCpu)
            {
                CpuData.Add(new data { value = a1, doc = DateTime.Now });
            }
            else
            {
                CpuData.Clear();
            }
            decimal aMem = GetMemPer();
            if (aMem > setMem)
            {
                MemData.Add(new data { value = aMem, doc = DateTime.Now });
            }
            else
            {
                MemData.Clear();
            }
            if (diskUsage > setDisk)
            {
                DiskData.Add(new data { value = diskUsage, doc = DateTime.Now });
            }
            else
            {
                DiskData.Clear();
            }

            GetRes();
        }
        /// <summary>
        /// Get Memory utilization percentage
        /// </summary>
        /// <returns></returns>
        static decimal GetMemPer()
        {
            var wmiObject = new ManagementObjectSearcher("select * from Win32_OperatingSystem");

            var memoryValues = wmiObject.Get().Cast<ManagementObject>().Select(mo => new
            {
                FreePhysicalMemory = Double.Parse(mo["FreePhysicalMemory"].ToString()),
                TotalVisibleMemorySize = Double.Parse(mo["TotalVisibleMemorySize"].ToString())
            }).FirstOrDefault();

            if (memoryValues != null)
            {
                var percent = ((memoryValues.TotalVisibleMemorySize - memoryValues.FreePhysicalMemory) / memoryValues.TotalVisibleMemorySize) * 100;
                return Convert.ToDecimal(percent);
            }
            else { return 0; }
        }

        /// <summary>
        /// Get System Events
        /// </summary>
        static void GetEvt()
        {
            Thread.Sleep(SetEVTTrackingTime);
            EventLog myLog = new EventLog();
            myLog.Log = "Application";

            aStr = myLog.Entries.Count.ToString();
            int a = 0;

            var entries = myLog.Entries.Cast<EventLogEntry>()
                         .Where(x => x.TimeGenerated >= DateTime.Now.AddMinutes(-1))
                         .OrderBy(x => x.TimeGenerated)
                         .Select(x => new
                         {
                             x.EventID,
                             x.InstanceId,
                             x.TimeGenerated,
                             x.MachineName,
                             x.Site,
                             x.Category,
                             x.Source,
                             x.Message
                         }).ToList();

            foreach (var entry in entries)
            {
                Dictionary<int, DateTime> dic = new Dictionary<int, DateTime>();
                dic.Add(entry.EventID, entry.TimeGenerated);
                bool isExist = keyValuePairs.Contains(dic);

                if (!isExist)
                {
                    keyValuePairs.Add(dic);

                    a++;
                    aStr += "********************************************************************" + Environment.NewLine;
                    aStr += "Event-Instance:" + entry.EventID.ToString() + "-" + entry.InstanceId.ToString() + Environment.NewLine;
                    aStr += "Datetime:" + entry.TimeGenerated.ToShortDateString() + "-" + entry.TimeGenerated.ToShortTimeString() + Environment.NewLine;
                    aStr += "Category:" + entry.Category + Environment.NewLine;
                    aStr += "Machine:" + entry.MachineName + Environment.NewLine;
                    aStr += "Message:" + entry.Message + Environment.NewLine;
                    aStr += "Source:" + entry.Source + Environment.NewLine;
                }
            }

        }

        /// <summary>
        /// Get Service Now Ticket raised by the local machine
        /// </summary>
        static async void GetSNT()
        {
            Thread.Sleep(SetSNTrackingTime);
            try
            {
                string aMachine = Environment.MachineName;
                //aMachine = "s-intelema-001";
                //aMachine = "IRULMATHI";
                aMachine = "V2P9M2";

                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Get, "https://wiprodemo4.service-now.com/api/now/table/incident?sysparm_query=cmdb_ci.name%3D" + aMachine + "&sysparm_limit=1");
                request.Headers.Add("X-UserToken", "87be18771b95f910807c20e0604bcb544d16d17010364ed60e8e02e6daae44bef2956ce3");
                request.Headers.Add("Authorization", "Basic cmFqaXYucGFkaGlAbHdwY29lLmNvbTpXaXByb0AxMjM0");
                request.Headers.Add("Cookie", "BIGipServerpool_wiprodemo4=629561098.34366.0000; JSESSIONID=DB2F201541916D323DBAE065786E6AFB; glide_node_id_for_js=89aa56208d868773dd4128c2cba5a5d48d1f56d757b8288ba394372d55cb041a; glide_session_store=4F84FBA81BFD8E106777631ABC4BCB97; glide_user=U0N2M18xOm5ydStlMTFXSTd5QUxiMjdueEwwckpJTnRKU0lHVW10Qm5IRHhmVTd1b2s9Om0wb1creU1FNWgyWXUySzNwemlndzF4UjZyRGkwTXZZRjB0MUhiem51UUU9; glide_user_route=glide.acd895053673453505bf5bcc2c835d68; glide_user_session=U0N2M18xOm5ydStlMTFXSTd5QUxiMjdueEwwckpJTnRKU0lHVW10Qm5IRHhmVTd1b2s9Om0wb1creU1FNWgyWXUySzNwemlndzF4UjZyRGkwTXZZRjB0MUhiem51UUU9");
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                string json = await response.Content.ReadAsStringAsync();

                var data = (JObject)JsonConvert.DeserializeObject(json);
                string data1 = data["result"][0]["short_description"].Value<string>();
                string sysid = data["result"][0]["sys_id"].Value<string>();
                //MessageBox.Show(data1);
                await new Form1().callLLMAsync(data1,false, sysid);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            GetSNT();
        }
        /// <summary>
        /// Create New Ticket if no resolution available in smartops for a non incident.
        /// </summary>
        /// <param name="shortDesc"></param>
        /// <param name="comment"></param>
        /// <returns></returns>
        static async Task<string> CreateTicket(string shortDesc, string comment)
        {
            string aTktNo = "";
            try
            {
                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Post, "https://wiprodemo4.service-now.com/api/now/v1/table/incident");
                request.Headers.Add("X-UserToken", "87be18771b95f910807c20e0604bcb544d16d17010364ed60e8e02e6daae44bef2956ce3");
                request.Headers.Add("Authorization", "Basic cmFqaXYucGFkaGlAbHdwY29lLmNvbTpXaXByb0AxMjM0");
                request.Headers.Add("Cookie", "BIGipServerpool_wiprodemo4=629561098.34366.0000; JSESSIONID=23A46115D0D9668BA264BC2FAAD48FB6; glide_node_id_for_js=89aa56208d868773dd4128c2cba5a5d48d1f56d757b8288ba394372d55cb041a; glide_session_store=BA110E111BB582506777631ABC4BCB63; glide_user_route=glide.acd895053673453505bf5bcc2c835d68");
                var content = new StringContent("{\r\n\"short_description\":\"" + shortDesc + "\",\r\n\"comments\":\"" + comment + "\"\r\n}", null, "application/json");
                request.Content = content;
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                string json = await response.Content.ReadAsStringAsync();

                var data = (JObject)JsonConvert.DeserializeObject(json);
                aTktNo = data["result"]["number"].Value<string>();
                string aSysId=data["result"]["sys_id"].Value<string>();
                aTktNo = aSysId;
                return aTktNo;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "";
            }
        }
        /// <summary>
        /// Update ticket if there is no resolution to an existing ticket
        /// </summary>
        /// <param name="sysid"></param>
        /// <param name="note"></param>
        static async void UpdateTicket(String sysid, string note)
        {

            try
            {
                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Put, "https://wiprodemo4.service-now.com/api/now/table/incident/" + sysid);
                request.Headers.Add("X-UserToken", "87be18771b95f910807c20e0604bcb544d16d17010364ed60e8e02e6daae44bef2956ce3");
                request.Headers.Add("Authorization", "Basic cmFqaXYucGFkaGlAbHdwY29lLmNvbTpXaXByb0AxMjM0");
                request.Headers.Add("Cookie", "BIGipServerpool_wiprodemo4=629561098.34366.0000; JSESSIONID=DB2F201541916D323DBAE065786E6AFB; glide_node_id_for_js=89aa56208d868773dd4128c2cba5a5d48d1f56d757b8288ba394372d55cb041a; glide_session_store=4F84FBA81BFD8E106777631ABC4BCB97; glide_user=U0N2M18xOm5ydStlMTFXSTd5QUxiMjdueEwwckpJTnRKU0lHVW10Qm5IRHhmVTd1b2s9Om0wb1creU1FNWgyWXUySzNwemlndzF4UjZyRGkwTXZZRjB0MUhiem51UUU9; glide_user_route=glide.acd895053673453505bf5bcc2c835d68; glide_user_session=U0N2M18xOm5ydStlMTFXSTd5QUxiMjdueEwwckpJTnRKU0lHVW10Qm5IRHhmVTd1b2s9Om0wb1creU1FNWgyWXUySzNwemlndzF4UjZyRGkwTXZZRjB0MUhiem51UUU9");
                var content = new StringContent("{\r\n\"work_notes\":\"" + note + "\"\r\n}", null, "application/json");
                request.Content = content;
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                string json = await response.Content.ReadAsStringAsync();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        /// <summary>
        /// execute script with file path
        /// </summary>
        /// <param name="pathToScript"></param>
        /// <returns></returns>
        public string ExecuteScript(string pathToScript)
        {
            var scriptArguments = "-ExecutionPolicy Bypass -File \"" + pathToScript + "\"";
            var processStartInfo = new ProcessStartInfo("powershell.exe", scriptArguments);
            processStartInfo.RedirectStandardOutput = true;
            processStartInfo.RedirectStandardError = true;

            var process = new Process();
            process.StartInfo = processStartInfo;
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            return output; // I am invoked using ProcessStartInfoClass!
        }
        /// <summary>
        /// execute powershell command
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        public string ExecuteCommand(string command)
        {
            var processStartInfo = new ProcessStartInfo();
            processStartInfo.FileName = "powershell.exe";
            processStartInfo.Arguments = $"-Command \"{command}\"";
            processStartInfo.UseShellExecute = false;
            processStartInfo.RedirectStandardOutput = true;

            var process = new Process();
            process.StartInfo = processStartInfo;
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            return output;
        }


        /// <summary>
        /// Nexthink Remote Action
        /// </summary>
        /// <param name="remoteactionId"></param>
        public async void CallNQL(string remoteactionId) {

            var client = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Post, "https://uxmlive.api.meta.nexthink.cloud/api/v1/token");
            request.Headers.Add("Authorization", "Basic N2EzZDEzNGUyMDEwNDY3OThkYTQzMjEwZmE4YmIwM2U6TlZINjJZYnNEUW5RZWUyd2daU2lzOElCYUszNmV5NkNlZzNaMU9vaEhaQXFyaHFsMmhEMmZDamFSRDM1cl90Rw==");
            var response = await client.SendAsync(request);
            response.EnsureSuccessStatusCode();
            string strAuth = await response.Content.ReadAsStringAsync();
            var AuthTok =  (JObject)JsonConvert.DeserializeObject(strAuth);
            string  token = AuthTok["access_token"].Value<string>();
            //getting token
            client = new HttpClient();
            request = new HttpRequestMessage(HttpMethod.Post, "https://uxmlive.api.meta.nexthink.cloud/api/v1/act/execute");
            request.Headers.Add("Authorization", "Bearer "+ token);
            var content = new StringContent("{\r\n  \"remoteActionId\": \""+ remoteactionId + "\",\r\n  \"devices\": [\r\n    \"30adce7c-657f-48d3-83e4-301e429475ce\"\r\n  ],\r\n  \"expiresInMinutes\": 60\r\n}", null, "application/json");
            request.Content = content;
            response = await client.SendAsync(request);
            response.EnsureSuccessStatusCode();
            string resJson= await response.Content.ReadAsStringAsync();
        }





        /// <summary>
        /// Call LLM API
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public async Task callLLMAsync(string query, bool RaiseTkt,string sysId="") {
            string aTktNo = sysId;//SysId
            if (RaiseTkt)
            {
                aTktNo = await CreateTicket(query, "created by agent");
            }
            var options = new RestClientOptions("http://127.0.0.1:8080")
            {
                MaxTimeout = -1,
            };
            var client = new RestClient(options);
            var request = new RestRequest("/apiAgent?query=" + query, Method.Post);
            request.AddHeader("Accept", "application/json");
            RestResponse response = await client.ExecuteAsync(request);

            var data = (JObject)JsonConvert.DeserializeObject(response.Content);
            string action = data["response"]["labels"][0].ToString();
            string aNote = data["response"]["sequence"].ToString();
            if (aTktNo!="")
            {
                //update ticket
                UpdateTicket(aTktNo, aNote);                
            }
            if (action == "Clear temp files") {
                tempClean();
            }
            if (action == "clean disk")
            {
                diskCleanup();
            }
            if (action == "Clear Browser settings")
            {
                browserClean();
            }
            if (action == "Restart")
            {
                RestartPC();
            }
            if (action == "delete Cache Cookies")
            {
                cacheClean();
            }
            //if (action == "create ticket")
            //{
            //    string shortDesc = data["response"]["short_desc"].Value<string>();
            //    string comment = data["response"]["comment"].Value<string>();
            //  sysId=  await CreateTicket(shortDesc, comment);
            //}
            if (action == "update ticket")
            {
                string note = data["response"]["note"].Value<string>();
                UpdateTicket(sysId, note);
            }
            if (action == "nexthink")
            {
                string remoteactionId = data["response"]["remoteactionId"].Value<string>();
                CallNQL(remoteactionId);
            }
            if (action == "Reinstall MS Teams")
            {
                string remoteactionId = "reinstall_microsoft_teams_windows";
                CallNQL(remoteactionId);
            }
        }



        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            CallNQL("reinstall_microsoft_teams_windows");
        }
    }

    public class data
    {
        public DateTime doc { get; set; }
        public decimal value { get; set; }
    }


}


//Restart
//Delete Temp Files
//Clear Cache Cookies

