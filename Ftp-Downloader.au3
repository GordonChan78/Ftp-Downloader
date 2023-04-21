#comments-start
Tool Name:
==========
Ftp Downloader

Arthor:
==========
Gordon Chan

Description:
==========
Ftp-Downloader is an automatic tool developed with AutoIt script(https://www.autoitscript.com/site/autoit/).
User may first config ftp server, file path, and email address, then start the downloading.
The tool will send notification email to the specified email address after the downloading is finished.

#comments-end

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

#include <ButtonConstants.au3>
#include <EditConstants.au3>

#include <file.au3>
#include <Array.au3>
#include <GuiStatusBar.au3>
#include <WinAPI.au3>

#include <Process.au3>

#include <FTPEx.au3>
#include <INet.au3>



;Set some option items
Opt("GUIOnEventMode", 0)
Opt ("TrayIconHide",0)
Opt("TrayMenuMode",1)
Opt("GUIDataSeparatorChar")

Global Const $Conf_File=@ScriptDir & "\Installer Download.ini"
Global Const $Log_File=@ScriptDir & "\Installer Download.log"
Global $Status_Bar
Global $FtpPath,$FileType,$LocalPath
Global $Local_File,$FTP_File_Size, $Local_File_Size

Global $oMyRet[2]
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

#Region ### UI Creation
$Main_Win=GUICreate("Installer Downloader", 400, 290, -1, -1, -1, 0)

GUICtrlCreateGroup("FTP Download", 5, 15, 380, 88)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlSetColor(-1, 0xff0000)
GUICtrlCreateLabel("From FTP Path:",15, 30, 95, 22 )
$FTP_Path = GUICtrlCreateEdit("", 110, 30, 250, 22,$ES_AUTOHSCROLL)
$Btn_FTP_Config = GUICtrlCreateButton("", 362, 30, 22,22,BitOR($BS_ICON, $BS_FLAT))
GUICtrlSetImage($Btn_FTP_Config, "shell32.dll", 24)
GUICtrlSetTip($Btn_FTP_Config,"Click this icon to configure FTP server","FTP Setting",1,1)

GUICtrlCreateLabel("Include Files:",15, 53, 95, 22 )
$File_Name = GUICtrlCreateEdit("*.*", 110, 53, 250, 22,$ES_AUTOHSCROLL)
GUICtrlCreateLabel("To Local Path:",15, 75, 95, 22 )
$Local_Path = GUICtrlCreateEdit("", 110, 75, 250, 22,$ES_AUTOHSCROLL)
GUICtrlCreateGroup("", -99, -99, 1, 1)


GUICtrlCreateGroup("Notication Email", 5,100, 383, 127)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlSetColor(-1, 0xff0000)
GUICtrlCreateLabel("Sent From:",15, 115, 90, 22 )
$Email_From = GUICtrlCreateInput("", 110, 115, 250, 22)
$Btn_Email_Config = GUICtrlCreateButton("", 362, 115, 22,22,BitOR($BS_ICON, $BS_FLAT))
GUICtrlSetImage($Btn_Email_Config, "shell32.dll", 24)
GUICtrlSetTip($Btn_Email_Config,"Click this icon to configure email server.","Email Setting",1,1)

GUICtrlCreateLabel("Send To:",15, 138, 90, 22 )
$Send_To = GUICtrlCreateInput("", 110, 138, 250, 22)
GUICtrlCreateLabel("CC To:",15, 161, 90, 22 )
$CC_To= GUICtrlCreateInput("", 110, 161, 250, 22)
GUICtrlCreateLabel("Email Subject:",15, 184, 90, 22 )
$Email_Subject = GUICtrlCreateInput("", 110, 184, 250, 22)
GUICtrlCreateLabel("Email Body:",15, 207, 90, 22 )
$Include_Log = GUICtrlCreateCheckbox("Include log file as email body", 110, 207, 250, 20)
GUICtrlSetState($Include_Log,$GUI_DISABLE)
GUICtrlSetState($Include_Log,$GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Btn_Start = GUICtrlCreateButton("Start Download", 100, 235, 150,22,$BS_FLAT)

$Status_Bar=_GUICtrlStatusBar_Create($Main_Win)
_GUICtrlStatusBar_SetText($Status_Bar,"Status: Ready for use...")


;Create the FTP Config dialog
$FTP_Config_Win = GUICreate("FTP Configuration", 380, 135, -1,-1,$WS_OVERLAPPED,$WS_EX_STATICEDGE,$Main_Win)
GUICtrlCreateLabel("FTP Server:", 12, 15, 80, 22)
$FTP_Server = GUICtrlCreateInput("", 100, 15, 130, 22)
GUICtrlCreateLabel("Port Nbr.:", 240, 15, 60, 22)
$FTP_Port = GUICtrlCreateInput("21", 300, 15, 60, 22)
GUICtrlCreateLabel("Username:", 12, 38, 80, 22)
$FTP_Username = GUICtrlCreateInput("", 100, 38, 130, 22)
GUICtrlCreateLabel("Password:", 240, 38, 60, 22)
$FTP_Password = GUICtrlCreateInput("", 300, 38, 60, 22)

$Btn_FTP_Save = GUICtrlCreateButton("Save", 150, 65, 40, 25, $BS_FLAT)
$Btn_FTP_Cancel = GUICtrlCreateButton("Exit", 190, 65, 40, 25, $BS_FLAT)


;Create the Email Config dialog
$Email_Config_Win = GUICreate("Email Configuration", 380, 135, -1,-1,$WS_OVERLAPPED,$WS_EX_STATICEDGE,$Main_Win)
GUICtrlCreateLabel("Mail Server:", 12, 15, 80, 22)
$Email_Server = GUICtrlCreateInput("", 100, 15, 130, 22)
GUICtrlCreateLabel("Port Nbr.:", 240, 15, 60, 22)
$Email_Port = GUICtrlCreateInput("", 300, 15, 60, 22)
GUICtrlCreateLabel("Username:", 12, 38, 80, 22)
$Email_Username = GUICtrlCreateInput("", 100, 38, 130, 22)
GUICtrlCreateLabel("Password:", 240, 38, 60, 22)
$Email_Password = GUICtrlCreateInput("", 300, 38, 60, 22)

$Btn_Email_Save = GUICtrlCreateButton("Save", 150, 65, 40, 25, $BS_FLAT)
$Btn_Email_Cancel = GUICtrlCreateButton("Exit", 190, 65, 40, 25, $BS_FLAT)

;Create the about dialog
$About_Win = GUICreate("About This Tool", 184, 140, -1,-1, BitOR($WS_CAPTION,$WS_POPUP,$WS_BORDER,$WS_CLIPSIBLINGS))
GUICtrlCreateGroup("", 8, 8, 169, 100)
GUICtrlCreateIcon("shell32.DLL", -19, 24, 32, 32, 32)
GUICtrlCreateLabel("Name: Installer Downloader", 64, 29, 150, 17, $WS_GROUP)
GUICtrlCreateLabel("Version: 0.1", 64, 50, 120, 17, $WS_GROUP)
GUICtrlCreateLabel("Publisher: xxx Inc. ", 24, 80, 150, 17, $WS_GROUP)
$Btn_About_OK = GUICtrlCreateButton("OK", 60, 111, 65, 17, $BS_FLAT)
#EndRegion ### UI Creation


GUISetState(@SW_SHOW,$Main_Win)
Load_Last_Data();

#EndRegion ### END Koda GUI section ###


#Region ### Event Loop
While 1
	$msg = GUIGetMsg()

	Select
		Case $msg=$GUI_EVENT_CLOSE
			$answer=MsgBox(4+32+262144, "Exit?", "Are you sure to exit Installer Downloader?")
			If $answer=6 Then

				Exit
			EndIf

	;Handle the button events in Main window
		Case $msg=$Btn_FTP_Config

			Load_FTP_Config()

			$win_pos=WinGetPos("Installer Downloader","")
			WinMove("FTP Configuration","",$win_pos[0]+400,$win_pos[1]+60)
			GUISetState(@SW_SHOW, $FTP_Config_Win)
			WinActivate("FTP Configuration","")

		Case $msg = $Btn_FTP_Save
			Save_FTP_Config()
			GUISetState(@SW_HIDE, $FTP_Config_Win)
			GUISetState(@SW_ENABLE,$Main_Win)
			WinActivate("Installer Downloader","")

		Case $msg = $Btn_FTP_Cancel
			GUISetState(@SW_HIDE, $FTP_Config_Win)
			GUISetState(@SW_ENABLE,$Main_Win)
			WinActivate("Installer Downloader","")


		Case $msg=$Btn_Email_Config
			Load_Email_Config()
			$win_pos=WinGetPos("Installer Downloader","")
			WinMove("Email Configuration","",$win_pos[0]+400,$win_pos[1]+60)
			GUISetState(@SW_SHOW, $Email_Config_Win)
			WinActivate("Email Configuration","")
		Case $msg = $Btn_Email_Save
			Save_Email_Config()
			GUISetState(@SW_HIDE, $Email_Config_Win)
			GUISetState(@SW_ENABLE,$Main_Win)
			WinActivate("Installer Downloader","")

		Case $msg = $Btn_Email_Cancel
			GUISetState(@SW_HIDE, $Email_Config_Win)
			GUISetState(@SW_ENABLE,$Main_Win)
			WinActivate("Installer Downloader","")
		Case $msg = $Btn_Start
			$result = Save_Last_Data();
			If $result = 1 Then
				Download_FTP_Files($FtpPath,$FileType,$LocalPath)
			EndIf


	EndSelect

WEnd
#EndRegion ### UI Creation


#Region ### User-defined functions



Func Load_FTP_Config()
	$Sec_Name = "FTP Configuration"
	GUICtrlSetData($FTP_Server,IniRead($Conf_File,$Sec_Name, "FTP Server",""))
	GUICtrlSetData($FTP_Port,IniRead($Conf_File,$Sec_Name, "Port Number",""))
	GUICtrlSetData($FTP_Username,IniRead($Conf_File,$Sec_Name, "Username",""))
	GUICtrlSetData($FTP_Password,IniRead($Conf_File,$Sec_Name, "Password",""))
EndFunc

Func Save_FTP_Config()
	$Sec_Name = "FTP Configuration"
	$Server = GUICtrlRead($FTP_Server)
	$Port = GUICtrlRead($FTP_Port)
	$Username = GUICtrlRead($FTP_Username)
	$Password = GUICtrlRead($FTP_Password)
	If ((StringLen($Server) = 0) Or (StringLen($Port) = 0) Or (StringLen($Username) = 0) Or (StringLen($Password) = 0) ) Then
		MsgBox(16, "Incomplete Configuration", "Please fill in all the fields with values.")
	Else
		IniWrite($Conf_File,$Sec_Name, "FTP Server",$Server)
		IniWrite($Conf_File,$Sec_Name, "Port Number",$Port)
		IniWrite($Conf_File,$Sec_Name, "Username",$Username)
		IniWrite($Conf_File,$Sec_Name, "Password",$Password)
	EndIf
EndFunc

Func Load_Email_Config()
	$Sec_Name = "Email Configuration"
	GUICtrlSetData($Email_Server,IniRead($Conf_File,$Sec_Name, "Email Server",""))
	GUICtrlSetData($Email_Port,IniRead($Conf_File,$Sec_Name, "Port Number",""))
	GUICtrlSetData($Email_Username,IniRead($Conf_File,$Sec_Name, "Username",""))
	GUICtrlSetData($Email_Password,IniRead($Conf_File,$Sec_Name, "Password",""))
EndFunc

Func Save_Email_Config()
	$Sec_Name = "Email Configuration"
	$Server = GUICtrlRead($Email_Server)
	$Port = GUICtrlRead($Email_Port)
	$Username = GUICtrlRead($Email_Username)
	$Password = GUICtrlRead($Email_Password)
	If ((StringLen($Server) = 0) Or (StringLen($Port) = 0) Or (StringLen($Username) = 0) Or (StringLen($Password) = 0) ) Then
		MsgBox(16, "Incomplete Configuration", "Please fill in all the fields with values.")
	Else
		IniWrite($Conf_File,$Sec_Name, "Email Server",$Server)
		IniWrite($Conf_File,$Sec_Name, "Port Number",$Port)
		IniWrite($Conf_File,$Sec_Name, "Username",$Username)
		IniWrite($Conf_File,$Sec_Name, "Password",$Password)

	EndIf
EndFunc

Func Load_Last_Data()
	$Sec_Name = "Last Data"
	GUICtrlSetData($FTP_Path,IniRead($Conf_File,$Sec_Name, "FTP Path",""))
	GUICtrlSetData($File_Name,IniRead($Conf_File,$Sec_Name, "File Type",""))
	GUICtrlSetData($Local_Path,IniRead($Conf_File,$Sec_Name, "Local Path",""))
	GUICtrlSetData($Email_From,IniRead($Conf_File,$Sec_Name, "Email From",""))
	GUICtrlSetData($Send_To,IniRead($Conf_File,$Sec_Name, "Send To",""))
	GUICtrlSetData($CC_To,IniRead($Conf_File,$Sec_Name, "CC To",""))
	GUICtrlSetData($Email_Subject,IniRead($Conf_File,$Sec_Name, "Email Subject",""))
EndFunc

Func Save_Last_Data()
	$Sec_Name = "Last Data"
	$FtpPath = GUICtrlRead($FTP_Path)
	$FileType = GUICtrlRead($File_Name)
	$LocalPath = GUICtrlRead($Local_Path)
	$EmailFrom = GUICtrlRead($Email_From)
	$SendTo = GUICtrlRead($Send_To)
	$CcTo = GUICtrlRead($CC_To)
	$EmailSubject = GUICtrlRead($Email_Subject)
	If  (StringLen($FtpPath) = 0 Or StringLen($FileType) = 0 Or StringLen($LocalPath) = 0) Then
		MsgBox(16, "Incomplete Input", "Please input all values in FTP Download section.")
		Return 0
	Else
		If (StringLen($EmailFrom) = 0 Or StringLen($SendTo) = 0 Or StringLen($CcTo) = 0 Or StringLen($EmailSubject) = 0) Then
			MsgBox(64, "Notification Email", "Email won't be sent because you have not input complete information." & @CRLF & "Click OK button to go on with downloading.")
		EndIf
		ProgressOn("Save Input","Saving the current input data...", "", -1 , @DesktopHeight / 2 - 100,16)
		IniWrite($Conf_File,$Sec_Name, "FTP Path",$FtpPath)
		IniWrite($Conf_File,$Sec_Name, "File Type",$FileType)
		IniWrite($Conf_File,$Sec_Name, "Local Path",$LocalPath)
		IniWrite($Conf_File,$Sec_Name, "Email From",$EmailFrom)
		IniWrite($Conf_File,$Sec_Name, "Send To",$SendTo)
		IniWrite($Conf_File,$Sec_Name, "CC To",$CcTo)
		IniWrite($Conf_File,$Sec_Name, "Email Subject",$EmailSubject)
		Return 1
	EndIf
EndFunc


Func Download_FTP_Files($FTP_Folder,$File_Name_Pattern,$Local_Folder)
	Dim $File_Array[20][2]
	$Sec_Name = "FTP Configuration"
	$F_Server = IniRead($Conf_File,$Sec_Name, "FTP Server","")
	$F_Port = IniRead($Conf_File,$Sec_Name, "Port Number","")
	$F_Username = IniRead($Conf_File,$Sec_Name, "Username","")
	$F_Password = IniRead($Conf_File,$Sec_Name, "Password","")

	$Email_Text = "Hello,"  & @CRLF & @CRLF

	$file = FileOpen($Log_File, 1+8)
	If ((StringLen($F_Server) = 0) Or (StringLen($F_Port) = 0) Or (StringLen($F_Username) = 0) Or (StringLen($F_Password) = 0) ) Then
		$Msg_Line = "The configuration of FTP server is not complete." & @CRLF
		MsgBox(16, "FTP Configuration", $Msg_Line)
		$Email_Text = @TAB & $Email_Text & $Msg_Line
	Else
	;	$FTP_Folder = GUICtrlRead($FTP_Path)
	;	$Local_Folder = GUICtrlRead($Local_Path)

		; Add slash for FTP path and Local path
		If (StringRight($FTP_Folder,1) <> "/") Then $FTP_Folder = $FTP_Folder & "/"
		If (StringRight($Local_Folder,1) <> "\" And StringRight($Local_Folder,1) <> "/") Then $Local_Folder = $Local_Folder & "\"

		$Open = _FTP_Open('FTP_Ctrl')
		$FTP_Conn = _FTP_Connect($Open, $F_Server, $F_Username, $F_Password,1)
	;	$FTP_Conn = _FTP_Connect($Open, $Server, $Username, $Password,0,$Port,1,0,0)
		_FTP_DirSetCurrent($FTP_Conn,$FTP_Folder)
		$Msg_Line = "********************** Date Time:" & Get_Date_Time() & " **********************" & @CRLF & _
					"Start to download files from FTP folder [" & $FTP_Folder & "] to local folder [" & $Local_Folder & "]:" & @CRLF & @CRLF
		FileWrite($file,$Msg_Line)
		$Email_Text = $Email_Text & $Msg_Line
		;Create local folder if it doesn't exist
		If (FileExists($Local_Folder) = 0) Then DirCreate($Local_Folder)
		If ($File_Name_Pattern = "*.*") Then   ;Download all files in batch
			_FTP_DirSetCurrent($FTP_Conn,$FTP_Folder)
			$File_Array = _FTP_ListToArray2D($FTP_Conn, 0)
			If ($File_Array[0][0] = 0) Then
				MsgBox(16, "No Files", "There is no files in the specified FTP folder" & @CRLF & $FTP_Folder)
				$Msg_Line = "Error: There is no file in the specified FTP folder: " & $FTP_Folder & @CRLF
				FileWrite($file,$Msg_Line)
				$Email_Text = @TAB & $Email_Text & $Msg_Line
			Else
				For $I = 3 To $File_Array[0][0]   ; Item 1 and 2 are directory names
					$File_Name = $File_Array[$I][0]
					$FTP_File_Size = $File_Array[$I][1]
					$Msg_Line = "  ==> Downloading file [name: " & $File_Name & "][size:" & $FTP_File_Size & " bytes]." & @CRLF
					FileWrite($file,$Msg_Line)
					$Email_Text = @TAB & $Email_Text & $Msg_Line
					$Local_File = $Local_Folder & $File_Name
					ProgressOn("FTP Download","File: " & $File_Name, "", -1 , @DesktopHeight / 2 - 100,16)
					_FTP_ProgressDownload($FTP_Conn,$Local_Folder & $File_Name, $File_Name,"_UpdateProgress")
					$Local_File_Size = FileGetSize($Local_Folder & $File_Name)
					If ($FTP_File_Size = $Local_File_Size) Then
						$Msg_Line = "    --> Downloaded the file successfully." & @CRLF & @CRLF
						FileWrite($file, $Msg_Line)
						$Email_Text = @TAB & @TAB & $Email_Text & $Msg_Line
					Else
						$Msg_Line = "    --> Failed to download the file." & @CRLF & @CRLF
						FileWrite($file, $Msg_Line)
						$Email_Text = @TAB & @TAB & $Email_Text & $Msg_Line
					EndIf
				Next
			EndIf
		Else   ;Download a single file based on the file name
			$File_Name = $File_Name_Pattern
			$FTP_File_Size = _FTP_FileGetSize($FTP_Conn, $File_Name)
			$Msg_Line = "  ==> Downloading file [name: " & $File_Name & "][size:" & $FTP_File_Size & " bytes]." & @CRLF
			FileWrite($file, $Msg_Line)
			$Email_Text = @TAB & $Email_Text & $Msg_Line
			ProgressOn("FTP Download","File: " & $File_Name, "", -1 , @DesktopHeight / 2 - 100,16)
			_FTP_ProgressDownload($FTP_Conn,$Local_Folder & $File_Name, $File_Name,"_UpdateProgress")
			$Local_File = $Local_Folder & $File_Name
			$Local_File_Size = FileGetSize($Local_File)
			If ($Local_File_Size = $FTP_File_Size) Then
				$Msg_Line = "    --> Downloaded the file successfully." & @CRLF & @CRLF
				FileWrite($file, $Msg_Line)
				$Email_Text = @TAB & @TAB & $Email_Text & $Msg_Line
			Else
				$Msg_Line = "    --> Failed to download the file." & @CRLF & @CRLF
				FileWrite($file, $Msg_Line)
				$Email_Text = @TAB & @TAB & $Email_Text & $Msg_Line
			EndIf
		EndIf

		$Email_Text = $Email_Text & @CRLF  & @CRLF & "Regards" & @CRLF & "Installer Downloader"
		FileClose($file)
		_FTP_Close($Open)
		ProgressOff()

		; Send notification email
		Local $objEmail = ObjCreate("CDO.Message")
		Local $i_Error = 0
		Local $i_Error_desciption = ""

		If ((GUICtrlRead($Send_To) <> "") And (GUICtrlRead($CC_To) <> "")) Then
			$Sec_Name = "Email Configuration"
			$SmtpServer = IniRead($Conf_File,$Sec_Name, "Email Server","")              ; address for the smtp-server to use - REQUIRED
			$FromName = "Installer Downloader"                      ; name from who the email was sent
			$FromAddress = GUICtrlRead($Email_From) ; address from where the mail should come
			$ToAddress = GUICtrlRead($Send_To)   ; destination address of the email - REQUIRED
			$Subject = GUICtrlRead($Email_Subject)                  ; subject from the email - can be anything you want it to be
			$Body = $Email_Text                              ; the messagebody from the mail - can be left blank but then you get a blank mail
			$AttachFiles = ""                       ; the file(s) you want to attach seperated with a ; (Semicolon) - leave blank if not needed
			$CcAddress = GUICtrlRead($CC_To)       ; address for cc - leave blank if not needed
			$BccAddress = ""     ; address for bcc - leave blank if not needed
			$Importance = "Normal"
			$Username = IniRead($Conf_File,$Sec_Name, "Username","")                  ; username for the account used from where the mail gets sent - REQUIRED
			$Password = IniRead($Conf_File,$Sec_Name, "Password","")                  ; password for the account used from where the mail gets sent - REQUIRED
			$IPPort = 25                            ; port used for sending the mail
			$ssl = 0                               ; enables/disables secure socket layer sending - put to 1 if using httpS
			;~ $IPPort=465                          ; GMAIL port used for sending the mail
			;~ $ssl=1                               ; GMAILenables/disables secure socket layer sending - put to 1 if using httpS

			$objEmail.From = '"' & $FromName & '" <' & $FromAddress & '>'
			$objEmail.To = $ToAddress

	;		MsgBox(64,"Debug","$SmtpServer=" & $SmtpServer & " $IPPort=" & $IPPort & " $Username=" & $Username & " $Password=" & $Password & @CRLF & " $FromName=" & $FromName & " $ToAddress=" & $ToAddress & " $CcAddress=" & $CcAddress & " $Subject=" & $Subject)
			$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $Username, $Password, $IPPort, $ssl)
			If @error Then
				MsgBox(0, "Email Error", "Failed to send email with error code:" & @error & "  Description:" & $rc)
			EndIf
		EndIf
	EndIf


EndFunc


Func _UpdateProgress($percent)
	GUICtrlSetData($Status_Bar,$percent)
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			_GUICtrlStatusBar_SetText($Status_Bar,"Cancelled the current downloading")
			Return -1 ; borts with -1, so you can exit you app afterwards
	EndSwitch
	$Local_File_Size = FileGetSize($Local_File)

	ProgressSet($percent,"Downloaded [" & $percent & "%] with " & $Local_File_Size & " / " & $FTP_File_Size & " bytes")

	_GUICtrlStatusBar_SetText($Status_Bar,"Downloaded [" & $percent & "%] with " & $Local_File_Size & " / " & $FTP_File_Size & " bytes")
	Return 1 ; Otherwise contine Download
 EndFunc

Func Get_Date_Time()
	Return (@YEAR & "-" & @MON & "-" & @MDAY & Chr(9) & @HOUR & ":" & @MIN )
	Exit
EndFunc


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject, $as_Body, $s_AttachFiles, $s_CcAddress, $s_BccAddress, $s_Importance, $s_Username, $s_Password, $IPPort, $ssl)


Local $objEmail = ObjCreate("CDO.Message")
    $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
    $objEmail.To = $s_ToAddress
    Local $i_Error = 0
    Local $i_Error_desciption = ""
    If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
    If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
    $objEmail.Subject = $s_Subject
    If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
        $objEmail.HTMLBody = $as_Body
    Else
        $objEmail.Textbody = $as_Body & @CRLF
    EndIf
    If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
        For $x = 1 To $S_Files2Attach[0]
            $S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
;~          ConsoleWrite('@@ Debug : $S_Files2Attach[$x] = ' & $S_Files2Attach[$x] & @LF & '>Error code: ' & @error & @LF) ;### Debug Console
            If FileExists($S_Files2Attach[$x]) Then
                ConsoleWrite('+> File attachment added: ' & $S_Files2Attach[$x] & @LF)
                $objEmail.AddAttachment($S_Files2Attach[$x])
            Else
                ConsoleWrite('!> File not found to attach: ' & $S_Files2Attach[$x] & @LF)
                SetError(1)
                Return 0
            EndIf
        Next
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    If Number($IPPort) = 0 then $IPPort = 25
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
    ;Update settings
    $objEmail.Configuration.Fields.Update
    ; Set Email Importance
    Switch $s_Importance
        Case "High"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "High"
        Case "Normal"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Normal"
        Case "Low"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Low"
    EndSwitch
    $objEmail.Fields.Update
    ; Sent the Message
    $objEmail.Send
    If @error Then
        SetError(2)
        Return $oMyRet[1]
    EndIf
    $objEmail=""
EndFunc   ;==>_INetSmtpMailCom


#EndRegion ### User-defined functions
