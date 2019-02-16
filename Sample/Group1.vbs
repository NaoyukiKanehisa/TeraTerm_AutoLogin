'==============================================================================================================
'' Tera Termのプログラム本体フォルダ名
TTPMACROPATH = "%PROGRAMFILES(X86)%\teraterm\ttpmacro.exe"

'' 対象Tera Termマクロのパス (絶対パス、または本スクリプトからの相対パスで指定)
TTLFILEPATH = "..\AutoLogin.ttl"

'' コマンドリストのパス (絶対パス、またはTTLファイルからの相対パスで指定)
'' ※ 使わない場合はブランク "" を設定する
COMMANDLIST = "command.list"

'' グループを指定
'' ※ 使わない場合はブランク "" を設定する
GROUP = "Group1"

'==============================================================================================================
Set WshShell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")
WORKDIR = Fso.GetParentFolderName(WScript.ScriptFullName)

TTPMACROPATH = WshShell.ExpandEnvironmentStrings(TTPMACROPATH)
TTPMACROPATHCHK = Fso.FileExists(TTPMACROPATH)

If TTPMACROPATHCHK = True Then
    '' UNCパスで本スクリプトが実行された場合
    If (InStr(1, WORKDIR, "\")) = 1 Then
        COMMAND = "cmd.exe /c, " &  """" & TTPMACROPATH & """"
        '' TTLのファイルパスが絶対パスの場合
        If (((InStr(1, TTLFILEPATH, "\")) = 1) Or ((InStr(1, TTLFILEPATH, ":")) = 2)) Then
            COMMAND = COMMAND & " " & """" & TTLFILEPATH & """"
        '' TTLファイルのパスが相対パスの場合、UNCパスに変換する
        Else
            COMMAND = COMMAND & " " & """" & (Fso.BuildPath(WORKDIR,TTLFILEPATH)) & """"
        End If
    '非UNCパスで本スクリプトが実行された場合'
    Else
        COMMAND = "cmd.exe /c, " & "cd " & """" & WORKDIR & """"
        COMMAND = COMMAND & " & " & """" & TTPMACROPATH & """"
        '' TTLファイルのパスが相対パスの場合、カレントディレクトリを示す%CD%をパスに付与する。
        If (((InStr(1, TTLFILEPATH, ":")) <> 2) And ((InStr(1, TTLFILEPATH, "\")) <> 1)) Then
            TTLFILEPATH = Fso.BuildPath("%CD%", TTLFILEPATH)
        End If
        COMMAND = COMMAND & " " & """" & TTLFILEPATH & """"
    End If
    COMMAND = COMMAND & " " & """" & (Fso.GetAbsolutePathName(COMMANDLIST)) & """" & " " & """" & GROUP & """"
    If (SCPSEND <> "") or (SCPSEND <> Null) Then
        COMMAND = COMMAND & " " & """" & SCPSEND & """"
    End If
    WshShell.Run(COMMAND),0,True
Else
    msgbox "Tera Termディレクトリ、またはttpmacro.exeが見つかりません。",4112,"エラー"
End If

Set WshShell = Nothing
Set Fso = Nothing