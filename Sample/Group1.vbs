'==============================================================================================================
'' Tera Term�̃v���O�����{�̃t�H���_��
TTPMACROPATH = "%PROGRAMFILES(X86)%\teraterm\ttpmacro.exe"

'' �Ώ�Tera Term�}�N���̃p�X (��΃p�X�A�܂��͖{�X�N���v�g����̑��΃p�X�Ŏw��)
TTLFILEPATH = "..\AutoLogin.ttl"

'' �R�}���h���X�g�̃J�X�^���p�X (��΃p�X�A�܂��͖{�X�N���v�g����̑��΃p�X�Ŏw��)
'' �� Tera Term�}�N�����̐ݒ�͖��������
'' �� �g��Ȃ��ꍇ�̓u�����N "" ��ݒ肷��
COMMANDLIST = "command.list"

'' �O���[�v���w��
'' �� �g��Ȃ��ꍇ�̓u�����N "" ��ݒ肷��
GROUP = "Group1"

'' ���O�̎����擾 (����: "on" | ���Ȃ�: "off" |  �N�����ɑI��: "select")
LOG_AUTOSTART = "on"

'' ���O�Ƀ^�C���X�^���v��ǉ� (����: "on" | ���Ȃ�: "off")
LOG_TIMESTAMP = "off"

'' ���O�t�H���_�̃J�X�^���ݒ�
'' �� Tera Term�}�N���̐ݒ���D�悳���B
LOG_DIR_PATH = "Sample\logs\%Y%m%d\$HOSTNAME"

'==============================================================================================================
Set WshShell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")
WORKDIR = Fso.GetParentFolderName(WScript.ScriptFullName)

TTPMACROPATH = WshShell.ExpandEnvironmentStrings(TTPMACROPATH)
TTPMACROPATHCHK = Fso.FileExists(TTPMACROPATH)

If TTPMACROPATHCHK = True Then
    'UNC�p�X�Ŗ{�X�N���v�g�����s���ꂽ�ꍇ
    If (InStr(1, WORKDIR, "\")) = 1 Then
        COMMAND = "cmd.exe /c, """ & TTPMACROPATH & """"
        '' TTL�̃t�@�C���p�X����΃p�X�̏ꍇ
        If (((InStr(1, TTLFILEPATH, "\")) = 1) Or ((InStr(1, TTLFILEPATH, ":")) = 2)) Then
            COMMAND = COMMAND & " """ & TTLFILEPATH & """"
        '' TTL�t�@�C���̃p�X�����΃p�X�̏ꍇ�AUNC�p�X�ɕϊ�����
        Else
            COMMAND = COMMAND & " """ & (Fso.BuildPath(WORKDIR,TTLFILEPATH)) & """"
        End If
    '��UNC�p�X�Ŗ{�X�N���v�g�����s���ꂽ�ꍇ
    Else
        COMMAND = "cmd.exe /c, cd /d """ & WORKDIR & """"
        COMMAND = COMMAND & " & """ & TTPMACROPATH & """"
        '' TTL�t�@�C���̃p�X�����΃p�X�̏ꍇ�A�J�����g�f�B���N�g��������%CD%���p�X�ɕt�^����B
        If (((InStr(1, TTLFILEPATH, ":")) <> 2) And ((InStr(1, TTLFILEPATH, "\")) <> 1)) Then
            TTLFILEPATH = Fso.BuildPath("%CD%", TTLFILEPATH)
        End If
        COMMAND = COMMAND & " """ & TTLFILEPATH & """"
    End If
    If ((LCase(LOG_AUTOSTART)) = "select") Then
        Result = MsgBox ("Tera Term���O���擾���܂����H",4099, "Tera Term���O�̎擾")
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
    MsgBox "Tera Term�f�B���N�g���A�܂���ttpmacro.exe��������܂���B",4112,"�G���["
End If

Set WshShell = Nothing
Set Fso = Nothing