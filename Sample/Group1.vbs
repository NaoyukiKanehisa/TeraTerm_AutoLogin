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

TTPMACROPATH = WshShell.ExpandEnvironmentStrings(TTPMACROPATH)
TTPMACROPATHCHK = Fso.FileExists(TTPMACROPATH)

If TTPMACROPATHCHK = True Then
    'UNCパスで本スクリプトが実行された場合
    If (InStr(1, WORKDIR, "\")) = 1 Then
        COMMAND = "cmd.exe /c, """ & TTPMACROPATH & """"
        '' TTLのファイルパスが絶対パスの場合
        If (((InStr(1, TTLFILEPATH, "\")) = 1) Or ((InStr(1, TTLFILEPATH, ":")) = 2)) Then
            COMMAND = COMMAND & " """ & TTLFILEPATH & """"
        '' TTLファイルのパスが相対パスの場合、UNCパスに変換する
        Else
            COMMAND = COMMAND & " """ & (Fso.BuildPath(WORKDIR,TTLFILEPATH)) & """"
        End If
    '非UNCパスで本スクリプトが実行された場合
    Else
        COMMAND = "cmd.exe /c, cd /d """ & WORKDIR & """"
        COMMAND = COMMAND & " & """ & TTPMACROPATH & """"
        '' TTLファイルのパスが相対パスの場合、カレントディレクトリを示す%CD%をパスに付与する。
        If (((InStr(1, TTLFILEPATH, ":")) <> 2) And ((InStr(1, TTLFILEPATH, "\")) <> 1)) Then
            TTLFILEPATH = Fso.BuildPath("%CD%", TTLFILEPATH)
        End If
        COMMAND = COMMAND & " """ & TTLFILEPATH & """"
    End If
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
    COMMAND = COMMAND & " """ & (Fso.GetAbsolutePathName(COMMANDLIST)) & """ """ & GROUP & """ """ & LOG_AUTOSTART & """ """ & LOG_TIMESTAMP & """ """ & LOG_DIR_PATH & """"
    WshShell.Run(COMMAND),0,True
Else
    MsgBox "Tera Termディレクトリ、またはttpmacro.exeが見つかりません。",4112,"エラー"
End If

Set WshShell = Nothing
Set Fso = Nothing