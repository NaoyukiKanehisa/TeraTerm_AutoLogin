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
TTPMACROPATH = Fso.GetAbsolutePathName((WshShell.ExpandEnvironmentStrings(TTPMACROPATH)))
TTPMACROPATHCHK = Fso.FileExists(TTPMACROPATH)
TTLFILEPATH = Fso.GetAbsolutePathName(TTLFILEPATH)
TTLFILEPATHCHK = Fso.FileExists(TTLFILEPATH)

If TTPMACROPATHCHK = True Then
    If TTLFILEPATHCHK = True Then
        COMMAND = "cmd.exe /c, """ & TTPMACROPATH & """ """ & TTLFILEPATH & """"
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

        IF (COMMANDLIST <> "") Then
            COMMANDLIST = Fso.GetAbsolutePathName(COMMANDLIST)
            COMMANDLISTPATHCHK = Fso.FileExists(COMMANDLIST)
            If COMMANDLISTPATHCHK = False Then
                MsgBox "�R�}���h���X�g��������܂���B" & vbCrLf &  vbCrLf & COMMANDLIST,4112,"�G���["
                WScript.Quit
            End If
        End If

        COMMAND = COMMAND & " """ & COMMANDLIST & """ """ & GROUP & """ """ & LOG_AUTOSTART & """ """ & LOG_TIMESTAMP & """ """ & LOG_DIR_PATH & """"
        WshShell.Run(COMMAND),0,True
    Else
        MsgBox "Tera Term�}�N����������܂���B" & vbCrLf &  vbCrLf & TTLFILEPATH,4112,"�G���["
        WScript.Quit
    End If
Else
    MsgBox "Tera Term�f�B���N�g���A�܂���ttpmacro.exe��������܂���B" & vbCrLf &  vbCrLf & TTPMACROPATH,4112,"�G���["
End If

Set WshShell = Nothing
Set Fso = Nothing