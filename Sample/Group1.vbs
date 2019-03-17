'==============================================================================================================
'' Tera Termのプログラム本体フォルダ名
TTPMACROPATH = "%PROGRAMFILES(X86)%\teraterm\ttpmacro.exe"

'' 対象Tera Termマクロのパス (絶対パス、または本スクリプトからの相対パスで指定)
TTLFILEPATH = "..\AutoLogin.ttl"

'' コマンドリストのカスタムパス (絶対パス、または本スクリプトからの相対パスで指定)
'' ※ Tera Termマクロ側の設定は無視される
'' ※ 使わない場合はブランク "" を設定する
COMMANDLIST = "command.list"

'' グループを指定
'' ※ 使わない場合はブランク "" を設定する
GROUP = "Group1"

'' ログの自動取得 (する: "on" | しない: "off" |  起動時に選択: "select")
LOG_AUTOSTART = "on"

'' ログにタイムスタンプを追加 (する: "on" | しない: "off")
LOG_TIMESTAMP = "off"

'' ログフォルダのカスタム設定
'' ※ Tera Termマクロの設定より優先される。
LOG_DIR_PATH = "Sample\logs\%Y%m%d\$HOSTNAME"

'==============================================================================================================
Set WshShell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")
WORKDIR = Fso.GetParentFolderName(WScript.ScriptFullName)
TTPMACROPATH = Fso.GetAbsolutePathName((WshShell.ExpandEnvironmentStrings(TTPMACROPATH)))
TTPMACROPATHCHK = Fso.FileExists(TTPMACROPATH)
TTLFILEPATH = Fso.GetAbsolutePathName(TTLFILEPATH)
TTLFILEPATHCHK = Fso.FileExists(TTLFILEPATH)

If TTPMACROPATHCHK = True Then
    If TTLFILEPATHCHK = True Then
        COMMAND = "cmd.exe /c, """ & TTPMACROPATH & """ """ & TTLFILEPATH & """"
        If ((LCase(LOG_AUTOSTART)) = "select") Then
            Result = MsgBox ("Tera Termログを取得しますか？",4099, "Tera Termログの取得")
            If (Result = 6) Then
                LOG_AUTOSTART = "on"
            ElseIf (Result = 7) Then
                LOG_AUTOSTART = "off"
            Else
                WScript.Quit
            End If
        End If

        IF (COMMANDLIST <> "") Then
            COMMANDLIST = Fso.GetAbsolutePathName(COMMANDLIST)
            COMMANDLISTPATHCHK = Fso.FileExists(COMMANDLIST)
            If COMMANDLISTPATHCHK = False Then
                MsgBox "コマンドリストが見つかりません。" & vbCrLf &  vbCrLf & COMMANDLIST,4112,"エラー"
                WScript.Quit
            End If
        End If

        COMMAND = COMMAND & " """ & COMMANDLIST & """ """ & GROUP & """ """ & LOG_AUTOSTART & """ """ & LOG_TIMESTAMP & """ """ & LOG_DIR_PATH & """"
        WshShell.Run(COMMAND),0,True
    Else
        MsgBox "Tera Termマクロが見つかりません。" & vbCrLf &  vbCrLf & TTLFILEPATH,4112,"エラー"
        WScript.Quit
    End If
Else
    MsgBox "Tera Termディレクトリ、またはttpmacro.exeが見つかりません。" & vbCrLf &  vbCrLf & TTPMACROPATH,4112,"エラー"
End If

Set WshShell = Nothing
Set Fso = Nothing