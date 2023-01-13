# Stormhen-Default-Email

Simple .VBS script that will make Stormhen the default e-mail client. Modified version of the original "thunderbirdportable.vbs" created by Ramesh Srinivasan for Winhelponline.com. All the credits (and thanks) to him.

I only changed the names to match the Stormhen executable name. If you want to have Thunderbird logo instead of Stormhen one, you need to modify these lines on the .vbs:  
  
If intOSBitness = 64 Then  
   sIconPath = sAppPath & "\stormhen-portable.exe"  
   sDLLPath = sAppPath & "\app\"  
ElseIf intOSBitness = 32 Then  
   sIconPath = sAppPath & "\stormhen-portable.exe"  
   sDLLPath = sAppPath & "\app\"  
  
For this one:  
  
If intOSBitness = 64 Then  
   sIconPath = sAppPath & "\app\thunderbird.exe"  
   sDLLPath = sAppPath & "\app\"  
ElseIf intOSBitness = 32 Then  
   sIconPath = sAppPath & "\app\thunderbird.exe"  
   sDLLPath = sAppPath & "\app\"  
