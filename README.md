# Panzer-Wireless-Gauge-VB6
 
  A FOSS Wireless Gauge VB6 WoW64 Widget for Windows Vista, 7, 8 and 10/11+. There will also be a version for Reactos and XP, watch this space for the link. Also tested and running well on Linux and Mac os/X using Wine.
 
 A current VB6/RC6 PSD program being worked upon now, added the Received Signal Strength Indicator (RSSI) value, so about 98% complete with some recent changes, awaiting: handling of dynamic WLAN access point changes, testing on Windows XP and Win7 32bit and some multi-monitor checking, completion of the CHM help file and the creation of the setup.exe. This Panzer widget is based upon the Yahoo/Konfabulator widget of the same visual design and very similar operation.

 Why VB6? Well, with a 64 bit, modern-language improvement upgrade on the way with 100% compatible TwinBasic coupled with support for transparent PNGs via RC/Cairo, VB6 code has an amazing future.

![vb6-logo-350](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/39e2c93f-40a5-4c47-8c23-d8ce7c747b10)

 I created as a variation of the previous gauges I had previously created for the World of Tanks and War Thunder 
 communities. The Panzer Wireless Gauge widget is an attractive dieselpunk VB6 widget for your desktop. 
 Functional and gorgeous at the same time. The graphics are my own, I took original inspiration from a clock face by Italo Fortana combining it with an aircraft gauge. It is all my code with some help from the chaps at VBForums (credits given). 
  
The Panzer Wireless Gauge VB6 is a useful utility displaying the wireless strength of all wi-fi network devices available to your system but it does so in a dieselpunk fashion on your desktop. This Widget is a moveable widget that you can move anywhere around the desktop as you require. 

The following are the declarations required for the various APIs used in this tool. This will goive you an indications as to how it obtains wifi information. You will have to dig into the code a little deeper to see the routine that obtains the data, it is too long to include here in full.

    Private Type GUID
        data1 As Long
        data2 As Integer
        data3 As Integer
        data4(7) As Byte
    End Type
    
    Private Type WLAN_INTERFACE_INFO
        ifGuid As GUID
        InterfaceDescription(511) As Byte
        IsState As Long
    End Type
    
    Private Type DOT11_SSID
        uSSIDLength As Long
        ucSSID(31) As Byte
    End Type
    
    Private Type WLAN_AVAILABLE_NETWORK
        strProfileName(511) As Byte
        dot11Ssid As DOT11_SSID
        dot11BssType As Long
        uNumberOfBssids As Long
        bNetworkConnectable As Long
        wlanNotConnectableReason As Long
        uNumberOfPhyTypes As Long
        dot11PhyTypes(7) As Long
        bMorePhyTypes As Long
        wlanSignalQuality As Long
        bSecurityEnabled As Long
        dot11DefaultAuthAlgorithm As Long
        dot11DefaultCipherAlgorithm As Long
        dwFlags As Long
        dwreserved As Long
    End Type
    
    Private Type AVAILABLE_NETWORK
        dot11Ssid As DOT11_SSID
        dot11BssType As Long
        uNumberOfBssids As Long
        bNetworkConnectable As Long
        wlanNotConnectableReason As Long
        uNumberOfPhyTypes As Long
        dot11PhyTypes(7) As Long
        bMorePhyTypes As Long
        wlanSignalQuality As Long
        bSecurityEnabled As Long
        dot11DefaultAuthAlgorithm As Long
        dot11DefaultCipherAlgorithm As Long
        dwFlags As Long
        dwreserved As Long
    End Type
    
    Private Type WLAN_INTERFACE_INFO_LIST
        dwNumberofItems As Long
        dwIndex As Long
        InterfaceInfo As WLAN_INTERFACE_INFO
    End Type
    
    Private Type WLAN_AVAILABLE_NETWORK_LIST
        dwNumberofItems As Long
        dwIndex As Long
        Network As WLAN_AVAILABLE_NETWORK
    End Type
    
    Private Type WLAN_BSS_LIST
        dwTotalSize As Long
        dwNumberofItems As Long
        wlanBssEntries As Long
    End Type
    
    Private Declare Function WlanOpenHandle Lib "wlanapi.dll" (ByVal dwClientVersion As Long, ByVal pdwReserved As Long, ByRef pdwNegotiaitedVersion As Long, ByRef phClientHandle As Long) As Long
    Private Declare Function WlanEnumInterfaces Lib "wlanapi.dll" (ByVal hClientHandle As Long, ByVal pReserved As Long, ppInterfaceList As Long) As Long
    Private Declare Function WlanGetAvailableNetworkList Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, ByVal dwFlags As Long, ByVal pReserved As Long, ppAvailableNetworkList As Long) As Long
    Private Declare Function WlanScan Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, pDot11Ssid As Long, pIeData As Long, reserved As Long) As Long
    Private Declare Function WlanGetNetworkBssList Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGui As GUID, ByVal pDot11Ssid As Long, ByVal dot11BssType As Long, ByVal bSecurityEnabled As Long, ByVal pReserved As Long, ppWlanBssList As Long) As Long
    Private Declare Function WlanGetProfile Lib "wlanapi.dll" (ByVal hClientHandle As Long, pInterfaceGuid As GUID, ByVal strProfileName As Long, ByVal pReserved As Long, pstrProfileXml As Long, pdwFlags As Long, pdwGrantedAccess As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Sub WlanFreeMemory Lib "wlanapi.dll" (ByVal pMemory As Long)
    'Private Declare Function lstrlenW& Lib "kernel32.dll" (ByVal lpszSrc&)
    
    Private udtList As WLAN_INTERFACE_INFO_LIST
    Private udtBSSList As WLAN_BSS_LIST
    Private ConIndex As Integer
    Private lHandle As Long
    Private lVersion As Long
    Private Connected As String
    Private bBuffer() As Byte

 Dig into the code, it is entirely FOSS, so help yourself!

 ![vb6PanzerWirelessPhoto1440x](https://github.com/yereverluvinunclebert/Panzer-Wireless-Gauge-VB6/assets/2788342/7e182a87-baf5-4244-bcb9-769487bdfb62)

 This widget can be increased in size, animation speed can be changed, 
 opacity/transparency may be set as to the users discretion. The widget can 
 also be made to hide for a pre-determined period.
 
![Default](https://github.com/yereverluvinunclebert/Panzer-Wireless-Gauge-VB6/assets/2788342/bd00835a-b72e-4d8c-929b-faa9291d749a)

 Right clicking will bring up a menu of options. Double-clicking on the widget will cause a personalised Windows application to 
 fire up. The first time you run it there will be no assigned function and so it 
 will state as such and then pop up the preferences so that you can enter the 
 command of your choice. The widget takes command line-style commands for 
 windows. Mouse hover over the widget and press CTRL+mousewheel up/down to resize. It works well on Windows XP 
 to Windows 11.
 
![panzergauge-help](https://github.com/yereverluvinunclebert/Panzer-Wireless-Gauge-VB6/assets/2788342/126d13a4-f726-40e7-b388-f9c2d3353f49)

 The Panzer CPU Gauge VB6 gauge is Beta-grade software, under development, not yet 
 ready to use on a production system - use at your own risk.

 This version was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS 
 project creating a WoW64 widget for the desktop. 
 
![Licence002](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/a24c5c45-5517-4423-938b-398f1d349d4c)

 It is open source to allow easy configuration, bug-fixing, enhancement and 
 community contribution towards free-and-useful VB6 utilities that can be created
 by anyone. The first step was the creation of this template program to form the 
 basis for the conversion of other desktop utilities or widgets. A future step 
 is new VB6 widgets with more functionality and then hopefully, conversion of 
 each to RADBasic/TwinBasic for future-proofing and 64bit-ness. 

![panzer-wireless0002](https://github.com/yereverluvinunclebert/Panzer-Wireless-Gauge-VB6/assets/2788342/7beb65e2-6574-43d2-92f8-a0a82f966617)


 This utility is one of a set of steampunk and dieselpunk widgets. That you can 
 find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery
 
 I do hope you enjoy using this utility and others. Your own software 
 enhancements and contributions will be gratefully received if you choose to 
 contribute.

![security](https://github.com/yereverluvinunclebert/Panzer-Wireless-Gauge-VB6/assets/2788342/fc73654d-5eab-4c06-9384-34c0a823399c)

 BUILD: The program runs without any Microsoft plugins.
 
 Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
 Framework 2.2 & Rubberduck 2.4.1, RichClient 6
 
 Links:
 
	https://www.vbrichclient.com/#/en/About/
	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Rocketdock https://punklabs.com/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	VBAdvance  
	Wireless APIs:
 		Doogle https://www.vbforums.com/showthread.php?633165-VB6-Wireless-Network-API-Exposed&highlight=wireless
		Coutttsj https://www.vbforums.com/showthread.php?881991-Simplified-WiFi-Scan

 
  Tested on :
 
	ReactOS 0.4.14 32bit on virtualBox    
	Windows 7 Professional 32bit on Intel    
	Windows 7 Ultimate 64bit on Intel    
	Windows 7 Professional 64bit on Intel    
	Windows XP SP3 32bit on Intel    
	Windows 10 Home 64bit on Intel    
	Windows 10 Home 64bit on AMD    
	Windows 11 64bit on Intel  
   
 CREDITS:
 
 I have really tried to maintain the credits as the project has progressed. If I 
 have made a mistake and left someone out then do forgive me. I will make amends 
 if anyone points out my mistake in leaving someone out.
 
 MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and 
 early 2000s. Thanks for VB6.
 
 Olaf Schmidt - This tool was built using the RichClient RC5 Cairo wrapper for 
 VB6. Specifically the components using transparency and reading images directly 
 from PSD. Thanks for the massive effort Olaf in creating Cairo counterparts for 
 all VB6 native controls and giving us access to advanced features on controls 
 such as transparency.
 
 Shuja Ali @ codeguru for his settings.ini code.
 
 ALLAPI.COM        For the registry reading code.
 
 Rxbagain on codeguru for his Open File common dialog code without a dependent 
 OCX - http://forums.codeguru.com/member.php?92278-rxbagain
 
 si_the_geek       for his special folder code
 
 Elroy on VB forums for the balloon tooltips.

 Doogle and Couttsj on the VB forums for their Wireless API code.
 
 Harry Whitfield for his quality testing, brain stimulation and being an 
 unwitting source of inspiration. 
 
 Dependencies:
 
 o A windows-alike o/s such as Windows XP, 7-11 or Apple Mac OSX 11. 
 
 o Microsoft VB6 IDE installed with its runtime components. The program runs 
 without any additional Microsoft OCX components, just the basic controls that 
 ship with VB6.  
 
 ![vb6-logo-200](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/bf00fa3d-f1d4-417b-bc50-9446f2c3e674)

 
 * Uses the latest version of the RC6 Cairo framework from Olaf Schmidt.
 
 During development the RC6 components need to be registered. These scripts are 
 used to register. Run each by double-clicking on them.
 
	RegisterRC6inPlace.vbs
	RegisterRC6WidgetsInPlace.vbs
 
 During runtime on the users system, the RC6 components are dynamically 
 referenced using modRC6regfree.bas which is compiled into the binary.	
 
 
 Requires a PzCPU Gauge folder in C:\Users\<user>\AppData\Roaming\ 
 eg: C:\Users\<user>\AppData\Roaming\PzCPU Gauge
 Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzCPU Gauge
 The above will be created automatically by the compiled program when run for the 
 first time.
 
 
o Krool's replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (slider) are replicated by the addition of one
dedicated OCX file that are shipped with this package.

During development only, this must be copied to C:\windows\syswow64 and should be registered.

- CCRSlider.ocx

Register this using regsvr32, ie. in a CMD window with administrator privileges.
	
	c:                          ! set device to boot drive with Windows
	cd \windows\syswow64s	    ! change default folder to syswow64
	regsvr32 CCRSlider.ocx	! register the ocx

This will allow the custom controls to be accessible to the VB6 IDE
at design time and the sliders will function as intended (if this ocx is
not registered correctly then the relevant controls will be replaced by picture boxes).

The above is only for development, for ordinary users, during runtime there is no need to do the above. The OCX will reside in the program folder. The program reference to this OCX is contained within the supplied resource file, Panzer Wireless Gauge.RES. The reference to this file is already compiled into the binary. As long as the OCX is in the same folder as the binary the program will run without the need to register the OCX manually.
 
 * OLEGuids.tlb
 
 This is a type library that defines types, object interfaces, and more specific 
 API definitions needed for COM interop / marshalling. It is only used at design 
 time (IDE). This is a Krool-modified version of the original .tlb from the 
 vbaccelerator website. The .tlb is compiled into the executable.
 For the compiled .exe this is NOT a dependency, only during design time.
 
 From the command line, copy the tlb to a central location (system32 or wow64 
 folder) and register it.
 
 COPY OLEGUIDS.TLB %SystemRoot%\System32\
 REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB
 
 In the VB6 IDE - project - references - browse - select the OLEGuids.tlb
 
![prefs-about](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/2349098d-f7df-4084-885e-383a58b87bac)


 * SETUP.EXE - The program is currently distributed using setup2go, a very useful 
 and comprehensive installer program that builds a .exe installer. Youll have to 
 find a copy of setup2go on the web as it is now abandonware. Contact me
 directly for a copy. The file "install PzCPU Gauge 0.1.0.s2g" is the configuration 
 file for setup2go. When you build it will report any errors in the build.
 
 * HELP.CHM - the program documentation is built using the NVU HTML editor and 
 compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and 
 the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
 the job admirably. The HTML files exist alongside the compiled CHM file in the 
 HELP folder.
 
  Project References:

	VisualBasic for Applications  
	VisualBasic Runtime Objects and Procedures  
	VisualBasic Objects and Procedures  
	OLE Automation  
	vbRichClient6  
 
 
 LICENCE AGREEMENTS:
 
 Copyright © 2023 Dean Beedell
 
 In addition to the GNU General Public Licence please be aware that you may use 
 any of my own imagery in your own creations but commercially only with my 
 permission. In all other non-commercial cases I require a credit to the 
 original artist using my name or one of my pseudonyms and a link to my site. 
 With regard to the commercial use of incorporated images, permission and a 
 licence would need to be obtained from the original owner and creator, ie. me.
 
![about](https://github.com/yereverluvinunclebert/Panzer-Wireless-Gauge-VB6/assets/2788342/5f2c76ca-9941-4390-8140-c0759d6f49e5)

![Panzer-CPU-Gauge-onDesktop](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/f2e06e77-77ee-46fb-840d-e14f96b4e0a5)





 
