'Registers Stormhen with Default Programs or Default Apps in Windows
'thunderbirdportable.vbs - created by Ramesh Srinivasan for Winhelponline.com
'StormHen-Default-Email.vbs - modified by Bunny-Head
'v1.0 23-July-2022 - Initial release. Tested on Thunderbird 102.0.3.
'StormHen-Default-Email Fork created 13-January-2023
'Suitable for all Windows versions, including Windows 10/11.
'Tutorial: https://www.winhelponline.com/blog/register-thunderbird-portable-with-default-apps/

Option Explicit
Dim sAction, sAppPath, sIconPath, objFile, sbaseKey, sMailKey, sAppDesc
Dim sDLLPath, sClassesKey, sOSBitness, sProfileDir, ArrKeys, regkey, intOSBitness
Dim sNewsKey, sCalKey, sExecPath
Dim WshShell : Set WshShell = CreateObject("WScript.Shell") 
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = oFSO.GetFile(WScript.ScriptFullName)
sAppPath = oFSO.GetParentFolderName(objFile)
sProfileDir = sAppPath & "\Data\profile"

'Quit if stormhen-portable.exe is missing in the current folder!
If Not oFSO.FileExists (sAppPath & "\stormhen-portable.exe") Then
   MsgBox "Please run this script from Stormhen folder. The script will now quit.", _
   vbOKOnly + vbInformation, "Register Stormhen with Default Apps"
   WScript.Quit
End If

'Determine if OS is 32-bit or 64-bit
intOSBitness = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth

'Get App and Icon paths
If intOSBitness = 64 Then
   sIconPath = sAppPath & "\stormhen-portable.exe"
   sDLLPath = sAppPath & "\app\"
ElseIf intOSBitness = 32 Then
   sIconPath = sAppPath & "\stormhen-portable.exe"
   sDLLPath = sAppPath & "\app\"
Else
   WScript.Quit
End If

sExecPath = sAppPath & "\stormhen-portable.exe"
sAppDesc = "Thunderbird is a full-featured email application. Thunderbird supports " & _
"IMAP and POP mail protocols, as well as HTML mail formatting. Built-in " & _
"junk mail controls, RSS capabilities, powerful quick search, " & _
"spell check as you type, global inbox, and advanced message filtering " & _
"round out Thunderbird's modern feature set."

If InStr(sExecPath, " ") > 0 Then sExecPath = """" & sExecPath & """"
If InStr(sIconPath, " ") > 0 Then sIconPath = """" & sIconPath & """"

sbaseKey = "HKCU\Software\"
sMailKey = sbaseKey & "Clients\Mail\Stormhen\"
sNewsKey = sbaseKey & "Clients\News\Stormhen\"
sCalKey = sbaseKey & "Clients\Calendar\Stormhen\"
sClassesKey = sbaseKey + "Classes\"

If WScript.Arguments.Count > 0 Then
   If UCase(Trim(WScript.Arguments(0))) = "-REG" Then Call RegisterThunderbirdPortable(sExecPath)
   If UCase(Trim(WScript.Arguments(0))) = "-UNREG" Then Call UnRegisterThunderbirdPortable
Else
   sAction = InputBox ("Type REGISTER to add Stormhen to Default Apps. Type UNREGISTER to remove.", _
   "Stormhen Registration", "REGISTER")
   If UCase(Trim(sAction)) = "REGISTER" Then Call RegisterThunderbirdPortable(sExecPath)
   If UCase(Trim(sAction)) = "UNREGISTER" Then Call UnRegisterThunderbirdPortable
End If


Sub RegisterThunderbirdPortable(sExecPath)
   
   'RegisteredApplications
   '----------------------------------------------------------------
   WshShell.RegWrite sbaseKey & "RegisteredApplications\Stormhen", _
   "Software\Clients\Mail\Stormhen\Capabilities", "REG_SZ"
   
   WshShell.RegWrite sbaseKey & "RegisteredApplications\Stormhen (News)", _
   "Software\Clients\News\Stormhen\Capabilities", "REG_SZ"
   
   WshShell.RegWrite sbaseKey & "RegisteredApplications\Stormhen (Calendar)", _
   "Software\Clients\Calendar\Stormhen\Capabilities", "REG_SZ"
   
   
   'ThunderbirdEML registration
   '----------------------------------------------------------------
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdEML2\", "Thunderbird Document", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdEML2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdEML2\FriendlyTypeName", "Thunderbird Document", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdEML2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdEML2\shell\open\command\", sExecPath & _
   " " & """" & "%1" & """", "REG_SZ"
   
   'Thunderbird Mailto registration
   '----------------------------------------------------------------
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mailto2\", "Thunderbird URL", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mailto2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mailto2\FriendlyTypeName", "Thunderbird URL", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mailto2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mailto2\shell\open\command\", sExecPath & _
   " -compose " & """" & "%1" & """", "REG_SZ"
   
   'Thunderbird News registration
   '----------------------------------------------------------------
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.news2\", "Thunderbird (News) URL", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.news2\FriendlyTypeName", "Thunderbird (News) URL", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.news2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.news2\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.news2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.news2\shell\open\command\", sExecPath & _
   " -mail " & """" & "%1" & """", "REG_SZ"
   
   'Thunderbird Mid registration
   '----------------------------------------------------------------
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mid2\", "Thunderbird URL", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mid2\FriendlyTypeName", "Thunderbird URL", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mid2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mid2\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mid2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.mid2\shell\open\command\", sExecPath & _
   " " & """" & "%1" & """", "REG_SZ"
   
   'Thunderbird Webcal registration
   '----------------------------------------------------------------
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.webcal2\", "Thunderbird (Calendar) URL", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.webcal2\FriendlyTypeName", _
   "Thunderbird (Calendar) URL", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.webcal2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.webcal2\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.webcal2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\Thunderbird.Url.webcal2\shell\open\command\", sExecPath & _
   " " & """" & "%1" & """", "REG_SZ"
   
   'Thunderbird ICS registration
   '----------------------------------------------------------------
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdICS2\", "Thunderbird (Calendar) Document", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdICS2\FriendlyTypeName", _
   "Thunderbird (Calendar) Document", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdICS2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdICS2\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdICS2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sbaseKey & "Classes\ThunderbirdICS2\shell\open\command\", sExecPath & _
   " " & """" & "%1" & """", "REG_SZ"
   
   
   
   'Default Apps Registration - Mail Capabilities
   '----------------------------------------------------------------
   WshShell.RegWrite sbaseKey & "Clients\Mail\", "Stormhen", "REG_SZ"
   WshShell.RegWrite sMailKey, "Stormhen", "REG_SZ"
   WshShell.RegWrite sMailKey & "Capabilities\ApplicationDescription", sAppDesc, "REG_SZ"
   WshShell.RegWrite sMailKey & "Capabilities\ApplicationIcon", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sMailKey & "Capabilities\ApplicationName", "Stormhen", "REG_SZ"
   WshShell.RegWrite sMailKey & "Capabilities\FileAssociations\.eml", "ThunderbirdEML2", "REG_SZ"
   WshShell.RegWrite sMailKey & "Capabilities\FileAssociations\.wdseml", "ThunderbirdEML2", "REG_SZ"
   WshShell.RegWrite sMailKey & "Capabilities\StartMenu\Mail", "Stormhen", "REG_SZ"
   WshShell.RegWrite sMailKey & "Capabilities\URLAssociations\mailto", "Thunderbird.Url.mailto2", "REG_SZ"
   WshShell.RegWrite sMailKey & "Capabilities\URLAssociations\mid", "Thunderbird.Url.mid2", "REG_SZ"
   
   WshShell.RegWrite sMailKey & "DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sMailKey & "shell\open\command\", sExecPath & " -mail", "REG_SZ"
   WshShell.RegWrite sMailKey & "shell\properties\", "Thunderbird &Options", "REG_SZ"
   WshShell.RegWrite sMailKey & "shell\properties\command\", sExecPath & " -options", "REG_SZ"
   WshShell.RegWrite sMailKey & "shell\safemode\", "Thunderbird &Safe Mode", "REG_SZ"   
   WshShell.RegWrite sMailKey & "shell\safemode\command\", sExecPath & " -safe-mode", "REG_SZ"
   
   WshShell.RegWrite sMailKey & "Protocols\mailto\", "Thunderbird URL", "REG_SZ"
   WshShell.RegWrite sMailKey & "Protocols\mailto\FriendlyTypeName", "Thunderbird URL", "REG_SZ"
   WshShell.RegWrite sMailKey & "Protocols\mailto\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sMailKey & "Protocols\mailto\Edit Flags", "2", "REG_DWORD"
   WshShell.RegWrite sMailKey & "Protocols\mailto\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sMailKey & "Protocols\mailto\shell\open\command\", sExecPath & _
   " -compose " & """" & "%1" & """", "REG_SZ"
   
   WshShell.RegWrite sMailKey & "Protocols\mid\", "Thunderbird URL", "REG_SZ"
   WshShell.RegWrite sMailKey & "Protocols\mid\FriendlyTypeName", "Thunderbird URL", "REG_SZ"
   WshShell.RegWrite sMailKey & "Protocols\mid\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sMailKey & "Protocols\mid\Edit Flags", "2", "REG_DWORD"
   WshShell.RegWrite sMailKey & "Protocols\mid\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sMailKey & "Protocols\mid\shell\open\command\", sExecPath & _
   " " & """" & "%1" & """", "REG_SZ"
   
   
   
   'Default Apps Registration - News Capabilities
   '----------------------------------------------------------------
   WshShell.RegWrite sNewsKey, "Stormhen", "REG_SZ"
   WshShell.RegWrite sNewsKey & "Capabilities\ApplicationDescription", sAppDesc, "REG_SZ"
   WshShell.RegWrite sNewsKey & "Capabilities\ApplicationIcon", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sNewsKey & "Capabilities\ApplicationName", "Stormhen (News)", "REG_SZ"
   WshShell.RegWrite sNewsKey & "Capabilities\URLAssociations\nntp", "Thunderbird.Url.news2", "REG_SZ"
   WshShell.RegWrite sNewsKey & "Capabilities\URLAssociations\news", "Thunderbird.Url.news2", "REG_SZ"
   WshShell.RegWrite sNewsKey & "Capabilities\URLAssociations\snews", "Thunderbird.Url.news2", "REG_SZ" 
   
   ArrKeys = Array("nntp", "news", "snews")
   For Each regkey In ArrKeys
      WshShell.RegWrite sNewsKey & "Protocols\" & regkey & "\", "Thunderbird (News) URL", "REG_SZ"
      WshShell.RegWrite sNewsKey & "Protocols\" & regkey & "\FriendlyTypeName", "Thunderbird (News) URL", "REG_SZ"
      WshShell.RegWrite sNewsKey & "Protocols\" & regkey & "\URL Protocol", "", "REG_SZ"
      WshShell.RegWrite sNewsKey & "Protocols\" & regkey & "\Edit Flags", "2", "REG_DWORD"
      WshShell.RegWrite sNewsKey & "Protocols\" & regkey & "\DefaultIcon\", sIconPath & ",0", "REG_SZ"
      WshShell.RegWrite sNewsKey & "Protocols\" & regkey & "\shell\open\command\", sExecPath & _
      " -mail " & """" & "%1" & """", "REG_SZ" 		
   Next
   
   
   'Default Apps Registration - Calendar Capabilities
   '----------------------------------------------------------------
   WshShell.RegWrite sCalKey & "DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sCalKey & "shell\open\command\", sExecPath & " -mail", "REG_SZ"
   WshShell.RegWrite sCalKey & "shell\properties\", "Thunderbird &Options", "REG_SZ"
   WshShell.RegWrite sCalKey & "shell\properties\command\", sExecPath & " -options", "REG_SZ"
   WshShell.RegWrite sCalKey & "shell\safemode\", "Thunderbird &Safe Mode", "REG_SZ"   
   WshShell.RegWrite sCalKey & "shell\safemode\command\", sExecPath & " -safe-mode", "REG_SZ"
   
   
   ArrKeys = Array(sMailKey, sNewsKey, sCalKey)
   For Each regkey In ArrKeys
      WshShell.RegWrite regkey, "Stormhen", "REG_SZ"
      WshShell.RegWrite regkey & "DefaultIcon", sIconPath & ",0", "REG_SZ"
      WshShell.RegWrite regkey & "DLLPath", sDLLPath & "mozMapi32.dll", "REG_SZ"
   Next
   
   WshShell.RegWrite sCalKey, "Stormhen", "REG_SZ"
   WshShell.RegWrite sCalKey & "Capabilities\ApplicationDescription", sAppDesc, "REG_SZ"
   WshShell.RegWrite sCalKey & "Capabilities\ApplicationIcon", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sCalKey & "Capabilities\ApplicationName", "Stormhen", "REG_SZ"  
   WshShell.RegWrite sCalKey & "Capabilities\FileAssociations\.ics", "ThunderbirdICS2", "REG_SZ"
   WshShell.RegWrite sCalKey & "Capabilities\URLAssociations\webcal", "Thunderbird.Url.webcal2", "REG_SZ"
   WshShell.RegWrite sCalKey & "Capabilities\URLAssociations\webcals", "Thunderbird.Url.webcal2", "REG_SZ"
   
   ArrKeys = Array("webcal", "webcals")
   
   For Each regkey In ArrKeys
      WshShell.RegWrite sCalKey & "Protocols\" & regkey & "\", "Stormhen (Calendar) URL", "REG_SZ"
      WshShell.RegWrite sCalKey & "Protocols\" & regkey & "\FriendlyTypeName", "Stormhen (Calendar) URL", "REG_SZ"
      WshShell.RegWrite sCalKey & "Protocols\" & regkey & "\URL Protocol", "", "REG_SZ"
      WshShell.RegWrite sCalKey & "Protocols\" & regkey & "\Edit Flags", "2", "REG_DWORD"
      WshShell.RegWrite sCalKey & "Protocols\" & regkey & "\DefaultIcon\", sIconPath & ",0", "REG_SZ"
      WshShell.RegWrite sCalKey & "Protocols\" & regkey & "\shell\open\command\", sExecPath & _
      " " & """" & "%1" & """", "REG_SZ"
   Next
   
   
   
   'Register DLLs/MAPI (for current user only)
   '----------------------------------------------------------------
   
   WshShell.RegWrite sClassesKey & "CLSID\{1814CEEB-49E2-407F-AF99-FA755A7D2607}\", _
   "PSFactoryBuffer", "REG_SZ"
   WshShell.RegWrite sClassesKey & "CLSID\{1814CEEB-49E2-407F-AF99-FA755A7D2607}\InProcServer32\", _
   sDLLPath & "AccessibleMarshal.dll", "REG_SZ"
   WshShell.RegWrite sClassesKey & "CLSID\{1814CEEB-49E2-407F-AF99-FA755A7D2607}\InProcServer32\ThreadingModel", _
   "Both", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "Interface\{0D68D6D0-D93D-4D08-A30D-F00DD1F45B24}\ProxyStubClsid32\", _
   "{1814CEEB-49E2-407F-AF99-FA755A7D2607}", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{0D68D6D0-D93D-4D08-A30D-F00DD1F45B24}\", _
   "ISimpleDOMDocument", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{0D68D6D0-D93D-4D08-A30D-F00DD1F45B24}\NumMethods\", _
   "9", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "Interface\{4E747BE5-2052-4265-8AF0-8ECAD7AAD1C0}\ProxyStubClsid32\", _
   "{1814CEEB-49E2-407F-AF99-FA755A7D2607}", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{4E747BE5-2052-4265-8AF0-8ECAD7AAD1C0}\", _
   "ISimpleDOMText", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{4E747BE5-2052-4265-8AF0-8ECAD7AAD1C0}\NumMethods\", _
   "8", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "Interface\{1814CEEB-49E2-407F-AF99-FA755A7D2607}\ProxyStubClsid32\", _
   "{1814CEEB-49E2-407F-AF99-FA755A7D2607}", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{1814CEEB-49E2-407F-AF99-FA755A7D2607}\", _
   "ISimpleDOMNode", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{1814CEEB-49E2-407F-AF99-FA755A7D2607}\NumMethods\", _
   "18", "REG_SZ"
   
   
   WshShell.RegWrite sClassesKey & "CLSID\{6E5AC413-1EF1-4985-87C2-F02E86E3ECDB}\InProcServer32\", _
   sDLLPath & "AccessibleHandler.dll", "REG_SZ"
   WshShell.RegWrite sClassesKey & "CLSID\{6E5AC413-1EF1-4985-87C2-F02E86E3ECDB}\InProcServer32\ThreadingModel", _
   "Apartment", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "CLSID\{127C0620-D3A7-45F1-8478-13D58249D68F}\", _
   "PSFactoryBuffer", "REG_SZ"
   WshShell.RegWrite sClassesKey & "CLSID\{127C0620-D3A7-45F1-8478-13D58249D68F}\InProcServer32\", _
   sDLLPath & "AccessibleHandler.dll", "REG_SZ"
   WshShell.RegWrite sClassesKey & "CLSID\{127C0620-D3A7-45F1-8478-13D58249D68F}\InProcServer32\ThreadingModel", _
   "Both", "REG_SZ"
   
   
   WshShell.RegWrite sClassesKey & "Interface\{127C0620-D3A7-45F1-8478-13D58249D68F}\ProxyStubClsid32\", _
   "{127C0620-D3A7-45F1-8478-13D58249D68F}", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{127C0620-D3A7-45F1-8478-13D58249D68F}\", _
   "IHandlerControl", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{127C0620-D3A7-45F1-8478-13D58249D68F}\NumMethods\", _
   "5", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{127C0620-D3A7-45F1-8478-13D58249D68F}\AsynchronousInterface\", _
   "{FE46DFA3-16B6-4780-93CC-4F6C174DC10D}", "REG_SZ"
   
   
   WshShell.RegWrite sClassesKey & "Interface\{FE46DFA3-16B6-4780-93CC-4F6C174DC10D}\", _
   "AsyncIHandlerControl", "REG_SZ"	
   WshShell.RegWrite sClassesKey & "Interface\{FE46DFA3-16B6-4780-93CC-4F6C174DC10D}\NumMethods\", _
   "7", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{FE46DFA3-16B6-4780-93CC-4F6C174DC10D}\SynchronousInterface\", _
   "{127C0620-D3A7-45F1-8478-13D58249D68F}", "REG_SZ"
   
   
   WshShell.RegWrite sClassesKey & "Interface\{0A8C4E6C-903F-41A9-B5D0-3520DB8F936B}\ProxyStubClsid32\", _
   "{127C0620-D3A7-45F1-8478-13D58249D68F}", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{0A8C4E6C-903F-41A9-B5D0-3520DB8F936B}\", _
   "IGeckoBackChannel", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{0A8C4E6C-903F-41A9-B5D0-3520DB8F936B}\NumMethods\", _
   "8", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "CLSID\{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}\InProcServer32\", _
   sDLLPath & "MapiProxy.dll", "REG_SZ"
   WshShell.RegWrite sClassesKey & "CLSID\{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}\InProcServer32\ThreadingModel", _
   "Both", "REG_SZ"
   WshShell.RegWrite sClassesKey & "CLSID\{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}\", _
   "PSFactoryBuffer", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "Interface\{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}\ProxyStubClsid32\", _
   "{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}\", _
   "nsIMapi", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Interface\{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}\NumMethods\", _
   "16", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "CLSID\{29F458BE-8866-11D5-A3DD-00B0D0F3BAA7}\LocalServer32\", _
   sDLLPath & "thunderbird.exe" & " -profile " & """" & sProfileDir & """" & " /MAPIStartup", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "CLSID\{29F458BE-8866-11D5-A3DD-00B0D0F3BAA7}\ProgID\",  _
   "MozillaMapi.1", "REG_SZ"
   WshShell.RegWrite sClassesKey & "CLSID\{29F458BE-8866-11D5-A3DD-00B0D0F3BAA7}\VersionIndependentProgID\", _
   "MozillaMAPI", "REG_SZ"
   
   WshShell.RegWrite sClassesKey & "MozillaMapi\", "MAPI", "REG_SZ"
   WshShell.RegWrite sClassesKey & "MozillaMapi\CLSID\", "{29F458BE-8866-11D5-A3DD-00B0D0F3BAA7}", "REG_SZ"
   WshShell.RegWrite sClassesKey & "MozillaMapi\CurVer\", "MozillaMapi.1", "REG_SZ"
   WshShell.RegWrite sClassesKey & "MozillaMapi.1\", "MAPI", "REG_SZ"
   WshShell.RegWrite sClassesKey & "MozillaMapi.1\CLSID\", "{29F458BE-8866-11D5-A3DD-00B0D0F3BAA7}", "REG_SZ"
   
   'App Paths registry key
   WshShell.RegWrite sbasekey & "Microsoft\Windows\CurrentVersion\App Paths\thunderbird.exe\",  _
   sExecPath, "REG_SZ"
   WshShell.RegWrite sbasekey & "Microsoft\Windows\CurrentVersion\App Paths\thunderbird.exe\Path", _
   sAppPath & "\", "REG_SZ"
   
   'Override the default app name by which the program appears in Default Apps  (*Optional*)
   '(i.e., -- "Thunderbird, Portable Edition" Vs. "Stormhen")
   'The official Thunderbird setup doesn't add this registry key.
   WshShell.RegWrite sClassesKey & "Thunderbird.Url.mailto2\Application\ApplicationIcon", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sClassesKey & "Thunderbird.Url.mailto2\Application\ApplicationName", "Stormhen", "REG_SZ"
   
   'Launch Default Programs or Default Apps after registering Stormhen   
   WshShell.Run "control /name Microsoft.DefaultPrograms /page pageDefaultProgram"
End Sub


Sub UnRegisterThunderbirdPortable
   
   sbaseKey = Replace(sbaseKey, "HKCU\", "HKEY_CURRENT_USER\")
   sClassesKey = Replace(sClassesKey, "HKCU\", "HKEY_CURRENT_USER\")
   
   Dim tmpfile: Set tmpfile = oFSO.OpenTextFile(sAppPath & "\unreg_tb_portable.reg", 2, True, 0)
   
   tmpfile.Writeline "REGEDIT4"
   tmpfile.writeblanklines 1
   tmpfile.Writeline "[" & sbaseKey & "RegisteredApplications" & "]"
   tmpfile.Writeline Chr(34) & "Stormhen" & Chr(34) & "=-"
   tmpfile.Writeline Chr(34) & "Stormhen (Calendar)" & Chr(34) & "=-"
   tmpfile.Writeline Chr(34) & "Stormhen (News)" & Chr(34) & "=-"
   tmpfile.writeblanklines 1
   tmpfile.Writeline "[" & sbaseKey & "Clients\Mail" & "]"
   tmpfile.Writeline "@=-"
   tmpfile.writeblanklines 1
   
   ArrKeys = Array ("Clients\Mail\Stormhen", _
   "Clients\News\Stormhen", _
   "Clients\Calendar\Stormhen", _
   "Microsoft\Windows\CurrentVersion\App Paths\thunderbird.exe")
   
   For Each regkey In ArrKeys
      tmpfile.Writeline "[-" & sbaseKey & regkey & "]"
      tmpfile.writeblanklines 1
   Next
   
   ArrKeys = Array("ThunderbirdEML2", _
   "Thunderbird.Url.mailto2", _
   "Thunderbird.Url.mid2", _
   "Thunderbird.Url.webcal2", _
   "MozillaMapi", _
   "MozillaMapi.1", _
   "CLSID\{1814CEEB-49E2-407F-AF99-FA755A7D2607}", _
   "CLSID\{6E5AC413-1EF1-4985-87C2-F02E86E3ECDB}", _
   "CLSID\{127C0620-D3A7-45F1-8478-13D58249D68F}", _
   "CLSID\{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}", _
   "CLSID\{29F458BE-8866-11D5-A3DD-00B0D0F3BAA7}", _
   "Interface\{0D68D6D0-D93D-4D08-A30D-F00DD1F45B24}", _
   "Interface\{4E747BE5-2052-4265-8AF0-8ECAD7AAD1C0}", _
   "Interface\{1814CEEB-49E2-407F-AF99-FA755A7D2607}", _
   "Interface\{127C0620-D3A7-45F1-8478-13D58249D68F}", _
   "Interface\{FE46DFA3-16B6-4780-93CC-4F6C174DC10D}", _
   "Interface\{0A8C4E6C-903F-41A9-B5D0-3520DB8F936B}", _
   "Interface\{6EDCD38E-8861-11D5-A3DD-00B0D0F3BAA7}")
   
   For Each regkey In ArrKeys
      tmpfile.Writeline "[-" & sClassesKey & regkey & "]"
      tmpfile.writeblanklines 1
   Next
   
   tmpfile.Close
   
   If oFSO.FileExists(sAppPath & "\unreg_tb_portable.reg") Then
      WshShell.Run "reg.exe import " & """" & sAppPath & "\unreg_tb_portable.reg" & """", 0, 1
      WScript.Sleep(1000)
      oFSO.DeleteFile sAppPath & "\unreg_tb_portable.reg"
   End If
   
   'Launch Default Apps after unregistering Stormhen
   WshShell.Run "control /name Microsoft.DefaultPrograms /page pageDefaultProgram"   
End Sub