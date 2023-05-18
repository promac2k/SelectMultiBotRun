#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Icon\Multibot.ico
#AutoIt3Wrapper_Outfile=SelectMultiBotRun.Exe
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=Made by Fliegerfaust, Edited for MultiBotRun by ProMac
#AutoIt3Wrapper_Res_Description=SelectMultiBotRun for MultiBotRun
#AutoIt3Wrapper_Res_Fileversion=1.0.6.0
#AutoIt3Wrapper_Res_LegalCopyright=Fliegerfaust, edited by ProMac
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****



#include <File.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <AutoItConstants.au3>
#include <ListBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <GuiButton.au3>
#include <TrayConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ComboConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <ProcessConstants.au3>
#include <WinAPIProc.au3>
#include <WinAPISys.au3>
#include <GuiMenu.au3>
#include <InetConstants.au3>
#include <Misc.au3>
#include <WinAPI.au3>
#include <EditConstants.au3>
#include <ColorConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <GuiStatusBar.au3>
#include <ProgressConstants.au3>
#include <SendMessage.au3>
#include <Date.au3>

Global $g_sBotFile = "multibot.run.exe"
Global $g_sBotFileAU3 = "multibot.run.au3"
Global $g_sVersion = "1.0.7"
Global $g_sDirProfiles = @MyDocumentsDir & "\Profiles.ini"
Global $g_hGui_Main, $g_hGui_Profile, $g_hGui_Emulator, $g_hGui_Instance, $g_hGui_Dir, $g_hGui_Parameter, $g_hGUI_AutoStart, $g_hGUI_Edit, $g_hListview_Main, $g_hLst_AutoStart, $g_hLog, $g_hProgress, $g_hBtn_Shortcut, $g_hBtn_AutoStart, $g_hContext_Main
Global $g_hListview_Instances, $g_hLblUpdateAvailable
Global $g_aGuiPos_Main
Global $g_sTypedProfile, $g_sSelectedEmulator
Global $g_sIniProfile, $g_sIniEmulator, $g_sIniInstance, $g_sIniDir, $g_sIniParameters
Global $g_iParameters = 7
Global Enum $eRun = 1000, $eEdit, $eDelete, $eNickname

If @OSArch = "X86" Then
	$Wow6432Node = ""
	$HKLM = "HKLM"
Else
	$Wow6432Node = "\Wow6432Node"
	$HKLM = "HKLM64"
EndIf

GUI_Main()

Func GUI_Main()

	$g_hGui_Main = GUICreate("SelectMultiBotRun", 258, 452, -1, -1)
	$g_hListview_Main = GUICtrlCreateListView("", 8, 24, 241, 305, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS), -1)
	_GUICtrlListView_InsertColumn($g_hListview_Main, 1, "Setup", 172)
	_GUICtrlListView_InsertColumn($g_hListview_Main, 2, "Bot Vers", 65)
	$g_hLblUpdateAvailable = GUICtrlCreateLabel("(Update available)", 8, 4, 200, 17)
	GUICtrlSetFont(-1, 7)
	GUICtrlSetColor(-1, $COLOR_RED)
	GUICtrlSetState(-1, $GUI_HIDE)
	$g_hBtn_Setup = GUICtrlCreateButton("New Setup", 8, 336, 243, 25, $WS_GROUP)
	GUICtrlSetTip(-1, "Use this Button to create a new Setup with your Profile, wished Emulator and Instance aswell as the Bot you want to use")
	$g_hBtn_Shortcut = GUICtrlCreateButton("Shortcut", 8, 368, 113, 25, $WS_GROUP)
	GUICtrlSetTip(-1, "Use this Button to create a new Shortcut for the Selected Setups. With this you can run your Setups with a single Click")
	$g_hBtn_AutoStart = GUICtrlCreateButton("Auto Start", 136, 368, 113, 25, $WS_GROUP)
	GUICtrlSetTip(-1, "use this Button if you want to start Setups automatically when the Computer boots up")
	$hMenu_Help = GUICtrlCreateMenu("&Help")
	$hMenu_HelpMsg = GUICtrlCreateMenuItem("Help", $hMenu_Help)
	$hMenu_ForumTopic = GUICtrlCreateMenuItem("Forum Topic", $hMenu_Help)
	$hMenu_Documents = GUICtrlCreateMenuItem("Profile Directory", $hMenu_Help)
	$hMenu_Startup = GUICtrlCreateMenuItem("Startup Directory", $hMenu_Help)
	$hMenu_Emulators = GUICtrlCreateMenu("&Emulators")
	$hMenu_BlueStacks5 = GUICtrlCreateMenuItem("BlueStacks5", $hMenu_Emulators)
	$hMenu_MEmu = GUICtrlCreateMenuItem("MEmu", $hMenu_Emulators)
	$hMenu_Nox = GUICtrlCreateMenuItem("Nox", $hMenu_Emulators)
	$hMenu_Update = GUICtrlCreateMenu("Updates")
	$hMenu_CheckForUpdate = GUICtrlCreateMenuItem("Check for Updates", $hMenu_Update)
	$hMenu_Misc = GUICtrlCreateMenu("Misc")
	$hMenu_Clear = GUICtrlCreateMenuItem("Clear Local Files", $hMenu_Misc)

	$g_hLog = _GUICtrlStatusBar_Create($g_hGui_Main)
	_GUICtrlStatusBar_SetParts($g_hLog, 2, 185)
	_GUICtrlStatusBar_SetText($g_hLog, "Version: " & $g_sVersion)
	$g_hProgress = GUICtrlCreateProgress(0, 0, -1, -1, $PBS_SMOOTH)
	$hProgress = GUICtrlGetHandle($g_hProgress)
	_GUICtrlStatusBar_EmbedControl($g_hLog, 1, $hProgress)

	$g_hContext_Main = _GUICtrlMenu_CreatePopup()
	_GUICtrlMenu_InsertMenuItem($g_hContext_Main, 0, "Run", $eRun)
	_GUICtrlMenu_InsertMenuItem($g_hContext_Main, 1, "Edit", $eEdit)
	_GUICtrlMenu_InsertMenuItem($g_hContext_Main, 2, "Delete", $eDelete)
	_GUICtrlMenu_InsertMenuItem($g_hContext_Main, 3, "Nickname", $eNickname)
	GUISetState(@SW_SHOW)
	GUIRegisterMsg($WM_CONTEXTMENU, "WM_CONTEXTMENU")
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

	UpdateList_Main()
	CheckUpdate()
	ChangeLog()

	If IniRead($g_sDirProfiles, "Options", "DisplayVersSent", "") = "" Then IniWrite($g_sDirProfiles, "Options", "DisplayVersSent", "1.0")

	While 1
		$aMsg = GUIGetMsg(1)
		Switch $aMsg[1]

			Case $g_hGui_Main
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						ExitLoop

					Case $hMenu_HelpMsg
						MsgBox($MB_OK, "Help", "To create a new Setup just press the New Setup Button and walk through the Guide!" & @CRLF & @CRLF & "To create a new Shortcut just press the New Shortcut Button and a Shortcut gets created on your Desktop!" & @CRLF & @CRLF & "Double Click an Item in the List to start the Bot with the highlighted Setup!" & @CRLF & @CRLF & "Right Click for a Context Menu." & @CRLF & @CRLF & "The Auto Updater will be downloaded and when you turn it off it will stay there but won't activate. When you delete this Tool make sure to Click on Misc and then Clear Local Files!", 0, $g_hGui_Main)
					Case $hMenu_ForumTopic
						ShellExecute("https://forum.multibot.run/index.php?threads/multibotrun-selectmultibotrun-v1-0-1.44/")
					Case $hMenu_Documents
						ShellExecute(@MyDocumentsDir)
					Case $hMenu_Startup
						ShellExecute(@StartupDir)
					Case $hMenu_BlueStacks5
						ShellExecute("https://cdn3.bluestacks.com/downloads/windows/nxt/5.11.100.1063/a2851e52720cc67bfe72ee23599fcaa0/FullInstaller/x64/BlueStacksFullInstaller_5.11.100.1063_amd64_native.exe")
					Case $hMenu_MEmu
						ShellExecute("https://mega.nz/#!RR5FmQYb!qmpHcqzq1s5f6PPJfPRXadYx2AoEUtekjSeZr8kcvl4")
					Case $hMenu_Nox
						ShellExecute("https://mega.nz/#!AIwyQIII!iGkk3ed9iUaWT-PpVfiFANUMeDaiRwUnvhIjVI66_iw")
					Case $hMenu_CheckForUpdate
						$sTempPath = _WinAPI_GetTempFileName(@TempDir)
						$hUpdateFile = InetGet("https://raw.githubusercontent.com/tehbank/SelectMultiBotRun/master/SelectMultiBotRun_Info.txt", $sTempPath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
						Do
							Sleep(250)
						Until InetGetInfo($hUpdateFile, $INET_DOWNLOADCOMPLETE)

						InetClose($hUpdateFile)
						$hGitVersion = IniRead($sTempPath, "General", "DisplayVers", "")
						$sGitVersion = StringStripWS($hGitVersion, 8)
						$Update = _VersionCompare($g_sVersion, $sGitVersion)

						Select
							Case $Update = -1
								_GUICtrlStatusBar_SetText($g_hLog, "Update found!")
								$msgUpdate = MsgBox($MB_YESNO, "Update", "New SelectMultiBotRun Update found" & @CRLF & "New: " & $sGitVersion & @CRLF & "Old: " & $g_sVersion & @CRLF & "Do you want to download it?", 0, $g_hGui_Main)
								If $msgUpdate = $IDYES Then
									UpdateSelect()
								EndIf
							Case $Update = 0
								_GUICtrlStatusBar_SetText($g_hLog, "Up to date!")
								MsgBox($MB_OK, "Update", "No new Update found (" & $g_sVersion & ")")
							Case $Update = 1
								_GUICtrlStatusBar_SetText($g_hLog, "Are you a magician?")
								MsgBox($MB_OK, "Update", "You are using a future Version (" & $g_sVersion & ")")
						EndSelect
						FileDelete($sTempPath)
					Case $hMenu_Clear
						$hMsgDelete = MsgBox($MB_YESNO, "Delete Local Files", "This will delete all SelectMultiBotRun Files (Profiles, Config and Auto Update) Do you want to proceed?", 0, $g_hGui_Main)
						If $hMsgDelete = 6 Then
							_GUICtrlStatusBar_SetText($g_hLog, "Deleting Files")
							FileDelete(@StartupDir & "\SelectMultiBotRunAutoUpdate.exe")
							FileDelete($g_sDirProfiles)
							UpdateList_Main()
							_GUICtrlStatusBar_SetText($g_hLog, "Done")
							MsgBox($MB_OK, "Delete Local Files", "Deleted all Files from your Computer!", 0, $g_hGui_Main)

						EndIf
					Case $g_hBtn_Setup
						Local $bSetupStopped = False

						_GUICtrlStatusBar_SetText($g_hLog, "Creating new Setup")
						WinSetOnTop($g_hGui_Main, "", $WINDOWS_ONTOP)

						Do
							GUISetState(@SW_DISABLE, $g_hGui_Main)
							$g_aGuiPos_Main = WinGetPos($g_hGui_Main)
							_GUICtrlStatusBar_SetText($g_hLog, "Select Profile")
							$bSetupStopped = GUI_Profile()
							If $bSetupStopped Then ExitLoop
							GUICtrlSetData($g_hProgress, 20)
							$bSetupStopped = GUI_Emulator()
							If $bSetupStopped Then ExitLoop
							GUICtrlSetData($g_hProgress, 40)
							$bSetupStopped = GUI_Instance()
							If $bSetupStopped Then ExitLoop
							GUICtrlSetData($g_hProgress, 60)
							$bSetupStopped = GUI_DIR()
							If $bSetupStopped Then ExitLoop
							GUICtrlSetData($g_hProgress, 80)
							$bSetupStopped = GUI_PARAMETER()
							If $bSetupStopped Then ExitLoop
							_GUICtrlStatusBar_SetText($g_hLog, "Setup Created!")
						Until 1

						WinSetOnTop($g_hGui_Main, "", $WINDOWS_NOONTOP)
						If $bSetupStopped Then _GUICtrlStatusBar_SetText($g_hLog, "Create Setup stopped")
						$bSetupStopped = False

						GUICtrlSetData($g_hProgress, 0)
						GUISetState(@SW_ENABLE, $g_hGui_Main)
						UpdateList_Main()
					Case $g_hBtn_AutoStart
						GUISetState(@SW_DISABLE, $g_hGui_Main)
						$g_aGuiPos_Main = WinGetPos($g_hGui_Main)
						GUI_AutoStart()
						GUISetState(@SW_ENABLE, $g_hGui_Main)
					Case $g_hBtn_Shortcut
						CreateShortcut()
				EndSwitch
		EndSwitch
	WEnd
EndFunc   ;==>GUI_Main

Func GUI_Profile()
	$g_hGui_Profile = GUICreate("Profile", 255, 167, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
	$hIpt_Profile = GUICtrlCreateInput("", 24, 72, 201, 21)
	$hBtn_Next = GUICtrlCreateButton("Next step", 72, 120, 97, 25, $WS_GROUP)
	GUICtrlSetState($hBtn_Next, $GUI_DISABLE)
	GUICtrlCreateLabel("Please type in the full Name of your Profile to continue", 24, 8, 204, 57)
	GUISetState()
	Local $bBtnEnabled = False
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				$g_sTypedProfile = GUICtrlRead($hIpt_Profile)
				IniDelete($g_sDirProfiles, $g_sTypedProfile)
				GUIDelete($g_hGui_Profile)
				Return -1
			Case $hBtn_Next
				$g_sTypedProfile = GUICtrlRead($hIpt_Profile)
				If $g_sTypedProfile = "" Then
					_GUICtrlStatusBar_SetText($g_hLog, "Profile cannot be empty!")
					ContinueLoop
				Else
					IniWrite($g_sDirProfiles, $g_sTypedProfile, "Profile", $g_sTypedProfile)
					IniWrite($g_sDirProfiles, $g_sTypedProfile, "BotVers", "")
					GUIDelete($g_hGui_Profile)
					_GUICtrlStatusBar_SetText($g_hLog, "Profile: " & $g_sTypedProfile)
					Return 0
				EndIf
		EndSwitch

		If $bBtnEnabled Then ContinueLoop
		If GUICtrlRead($hIpt_Profile) Then
			ContinueLoop
		Else
			GUICtrlSetState($hBtn_Next, $GUI_ENABLE)
			$bBtnEnabled = True
		EndIf
	WEnd
EndFunc   ;==>GUI_Profile

Func GUI_Emulator()
	$g_hGui_Emulator = GUICreate("Emulator", 258, 167, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
	$hCmb_Emulator = GUICtrlCreateCombo("BlueStacks5", 24, 72, 201, 21, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
	GUICtrlSetData(-1, "BlueStacks5|MEmu|Nox")
	$hBtn_Next = GUICtrlCreateButton("Next step", 72, 120, 97, 25, $WS_GROUP)
	GUICtrlCreateLabel("Please select your Emulator", 24, 8, 204, 57)
	GUISetState()

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				IniDelete($g_sDirProfiles, $g_sTypedProfile)
				GUIDelete($g_hGui_Emulator)
				Return -1
			Case $hBtn_Next
				$g_sSelectedEmulator = GUICtrlRead($hCmb_Emulator)
				If IsAndroidInstalled($g_sSelectedEmulator) Then
					IniWrite($g_sDirProfiles, $g_sTypedProfile, "Emulator", $g_sSelectedEmulator)
					GUIDelete($g_hGui_Emulator)
					_GUICtrlStatusBar_SetText($g_hLog, "Emulator: " & $g_sSelectedEmulator)
					Return 0
				Else
					$msgEmulator = MsgBox($MB_YESNO, "Error", "Sorry Chief!" & @CRLF & "Couldn't find " & $g_sSelectedEmulator & " installed on your Computer. Did you chose the wrong Emulator ? If you are sure you got it installed please click 'Yes'" & @CRLF & @CRLF & "Do you want to continue?", 0, $g_hGui_Emulator)
					If $msgEmulator = $IDYes Then
						IniWrite($g_sDirProfiles, $g_sTypedProfile, "Emulator", $g_sSelectedEmulator)
						GUIDelete($g_hGui_Emulator)
						_GUICtrlStatusBar_SetText($g_hLog, "Emulator: " & $g_sSelectedEmulator)
						Return 0
					EndIf
				EndIf
		EndSwitch
	WEnd
EndFunc   ;==>GUI_Emulator

Func GUI_Instance()
	Local $hLbl_Instance = 0

	$g_hGui_Instance = GUICreate("Instance", 258, 167, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
	$hIpt_Instance = GUICtrlCreateInput("", 24, 72, 201, 21)
	$hBtn_Next = GUICtrlCreateButton("Next step", 72, 120, 97, 25, $WS_GROUP)
	$hLbl_Instance = GUICtrlCreateLabel("Please type in the Instance Name you want to use", 24, 8, 204, 57)
	GUISetState(@SW_HIDE, $g_hGui_Instance)

	Switch $g_sSelectedEmulator
		Case "BlueStacks", "BlueStacks5"
			Return
		Case "Bluestacks5"
			GUISetState(@SW_SHOW, $g_hGui_Instance)
			GUICtrlSetData($hLbl_Instance, "Please type in your BlueStacks5 Instance Name! Example: Pie64 , Pie64_1, Pie64_2, etc")
			GUICtrlSetData($hIpt_Instance, "Pie64_")
		Case "MEmu"
			GUISetState(@SW_SHOW, $g_hGui_Instance)
			GUICtrlSetData($hLbl_Instance, "Please type in your MEmu Instance Name! Example: MEmu , MEmu_1, MEmu_2, etc")
			GUICtrlSetData($hIpt_Instance, "MEmu_")
		Case "Nox"
			GUISetState(@SW_SHOW, $g_hGui_Instance)
			GUICtrlSetData($hLbl_Instance, "Please type in your Nox Instance Name! Example: nox , nox_1, nox_2, etc")
			GUICtrlSetData($hIpt_Instance, "nox_")
	EndSwitch

	While 1

		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				GUIDelete($g_hGui_Instance)
				IniDelete($g_sDirProfiles, $g_sTypedProfile)
				Return -1
			Case $hBtn_Next
				$Inst = GUICtrlRead($hIpt_Instance)
				$Instances = LaunchConsole(GetInstanceMgrPath($g_sSelectedEmulator), "list vms", 1000)
				Switch $g_sSelectedEmulator
					Case "BlueStacks5"
						$Instance = StringRegExp($Instances, "(?i)" & "Pie64" & "(?:[_][0-9])?", 3)
					Case "iTools"
						$Instance = StringRegExp($Instances, "(?)iToolsVM(?:[_][0-9][0-9])?", 3)
					Case Else
						$Instance = StringRegExp($Instances, "(?i)" & $g_sSelectedEmulator & "(?:[_][0-9])?", 3)
				EndSwitch
				_ArrayUnique($Instance, 0, 0, 0, 0)

				If _ArraySearch($Instance, $Inst, 0, 0, 1) = -1 And $Instance = 1 Then
					MsgBox($MB_OK, "Error", "Couldn't find any Instances for " & $g_sSelectedEmulator & "." & " There are only two reasons why." & @CRLF & "#1: You deleted all Instances" & @CRLF & "#2: You don't have the Emulator installed and still pressed YES on the Pop Up before :(", 0, $g_hGui_Instance)
					GUIDelete($g_hGui_Instance)
					IniDelete($g_sDirProfiles, $g_sTypedProfile)
					Return -1

				ElseIf _ArraySearch($Instance, $Inst, 0, 0, 1) = -1 Then
					$Msg2 = MsgBox($MB_YESNO, "Typo ?", "Couldn't find the Instance Name you typed in. Please check your Instances once again and retype it ( Also check the case sensitivity)" & @CRLF & "Here is a list of Instances I could find on your PC:" & @CRLF & @CRLF & _ArrayToString($Instance, @CRLF) & @CRLF & @CRLF & 'If you are sure that you got the Instance right but this Message keeps coming then press "Yes" to continue!' & @CRLF & @CRLF & "Do you want to continue?", 0, $g_hGui_Instance)
					If $Msg2 = $IDYES Then
						IniWrite($g_sDirProfiles, $g_sTypedProfile, "Instance", $Inst)
						GUIDelete($g_hGui_Instance)
						_GUICtrlStatusBar_SetText($g_hLog, "Instance: " & $Inst)
						Return 0
					EndIf
				Else
					IniWrite($g_sDirProfiles, $g_sTypedProfile, "Instance", $Inst)
					GUIDelete($g_hGui_Instance)
					_GUICtrlStatusBar_SetText($g_hLog, "Instance: " & $Inst)
					Return 0
				EndIf

		EndSwitch
	WEnd
EndFunc   ;==>GUI_Instance

Func GUI_DIR()
	$g_hGui_Dir = GUICreate("Directory", 258, 167, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
	$hBtn_Folder = GUICtrlCreateButton("Choose Folder", 24, 72, 201, 21)
	$hBtn_Finish = GUICtrlCreateButton("Next", 72, 120, 97, 25, $WS_GROUP)
	GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlCreateLabel("Please select the MultiBot Folder where the multibot.run.exe or .au3 is located at", 24, 8, 204, 57)
	GUISetState()

	While 1

		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				IniDelete($g_sDirProfiles, $g_sTypedProfile)
				GUIDelete($g_hGui_Dir)
				Return -1

			Case $hBtn_Folder
				WinSetOnTop($g_hGui_Main, "", $WINDOWS_NOONTOP)
				Local $sFileSelectFolder = FileSelectFolder("Select your MultiBot Folder", "")
				If @error Then
					;GUICtrlSetData($Lbl_Log, "File Selection aborted.")
				Else
					_GUICtrlStatusBar_SetText($g_hLog, "Dir: " & $sFileSelectFolder)
				EndIf
				WinSetOnTop($g_hGui_Main, "", $WINDOWS_ONTOP)


				If $sFileSelectFolder = "" Then
					ContinueLoop
				ElseIf FileExists($sFileSelectFolder & "\" & $g_sBotFile) = 0 And FileExists($sFileSelectFolder & "\" & $g_sBotFileAU3) = 0 Then
					MsgBox($MB_OK, "Error", "Looks like there is no runable multibot file in the Folder? Did you select the right folder or is in the Folder the multibot.run.exe or multibot.run.au3 renamed? Please select another Folder or rename Files!", 0, $g_hGui_Dir)
					ContinueLoop
				Else
					GUICtrlSetData($g_hProgress, 100)
					GUICtrlSetState($hBtn_Finish, $GUI_ENABLE)
				EndIf

			Case $hBtn_Finish

				IniWrite($g_sDirProfiles, $g_sTypedProfile, "Dir", $sFileSelectFolder)
				GUIDelete($g_hGui_Dir)
				ExitLoop
		EndSwitch
	WEnd
EndFunc   ;==>GUI_DIR

Func GUI_PARAMETER()
	Local $iEndResult

	$g_hGui_Parameter = GUICreate("Special Parameter", 258, 190, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
	GUICtrlCreateLabel("Check Options if you want to run the Bot with special Parameters. These are not mandatory!", 24, 8, 204, 57)
	$hChk_NoWatchdog = GUICtrlCreateCheckbox("No Watchdog", 8, 55)
	GUICtrlSetTip(-1, "Check this to run the Bot without Watchdog")
	$hChk_StartBotDocked = GUICtrlCreateCheckbox("Dock Bot on Start", 8, 75)
	GUICtrlSetTip(-1, "Auto Dock the Bot Window to the Android when the Bot launches")
	$hChk_StartBotDockedAndShrinked = GUICtrlCreateCheckbox("Dock & Shrink Bot on Start", 8, 95)
	GUICtrlSetTip(-1, "Auto Dock and Shrink the Bot Window to the Android when the Bot launches")
	$hChk_DpiAwarness = GUICtrlCreateCheckbox("DPI Awareness", 160, 55)
	GUICtrlSetTip(-1, "Launch the Bot in DPI Awareness Mode")
	$hChkDebugMode = GUICtrlCreateCheckbox("Debug Mode", 160, 75)
	GUICtrlSetTip(-1, "Launch the Bot with Debug Mode enabled")
	$hChk_MiniGUIMode = GUICtrlCreateCheckbox("Mini GUI Mode", 160, 95)
	GUICtrlSetTip(-1, "Launch the Bot in Mini GUI Mode")
	$hChk_HideAndroid = GUICtrlCreateCheckbox("Hide Android", 8, 115)
	GUICtrlSetTip(-1, "Hide the Android Emulator Window")

	$hBtn_Finish = GUICtrlCreateButton("Finish", 72, 160, 97, 25, $WS_GROUP)
	GUISetState()

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				IniDelete($g_sDirProfiles, $g_sTypedProfile)
				GUIDelete($g_hGui_Parameter)
				Return -1
			Case $hBtn_Finish
				$iEndResult &= GUICtrlRead($hChk_NoWatchdog) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_StartBotDocked) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_StartBotDockedAndShrinked) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_DpiAwarness) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChkDebugMode) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_MiniGUIMode) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_HideAndroid) = $GUI_CHECKED ? 1 : 0
				IniWrite($g_sDirProfiles, $g_sTypedProfile, "Parameters", $iEndResult)
				GUIDelete($g_hGui_Parameter)
				ExitLoop
		EndSwitch
	WEnd
EndFunc   ;==>GUI_PARAMETER



Func GUI_Edit()

	Local $iEndResult
	$aLstbx_GetSelTxt = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
	$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $aLstbx_GetSelTxt[1])
	ReadIni($sLstbx_SelItem)
	$sSelectedFolder = $g_sIniDir
	$g_hGUI_Edit = GUICreate("Edit INI", 258, 260, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
	$hIpt_Profile = GUICtrlCreateInput($g_sIniProfile, 112, 8, 137, 21)
	$hCmb_Emulator = GUICtrlCreateCombo($g_sIniEmulator, 112, 40, 137, 21, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
	$hIpt_Instance = GUICtrlCreateInput($g_sIniInstance, 112, 72, 137, 21)
	$hBtn_Folder = GUICtrlCreateButton("Open Dialog", 112, 105, 137, 23, $WS_GROUP)
	GUICtrlCreateLabel("Profile Name:", 8, 8, 95, 17)
	GUICtrlCreateLabel("Emulator:", 8, 40, 95, 17)
	GUICtrlCreateLabel("Instance:", 8, 72, 95, 17)
	GUICtrlCreateLabel("Directory:", 8, 104, 95, 17)

	$hChk_NoWatchdog = GUICtrlCreateCheckbox("No Watchdog", 8, 135)
	GUICtrlSetTip(-1, "Check this to run the Bot without Watchdog")
	$hChk_StartBotDocked = GUICtrlCreateCheckbox("Dock Bot on Start", 8, 155)
	GUICtrlSetTip(-1, "Auto Dock the Bot Window to the Android when the Bot launches")
	$hChk_StartBotDockedAndShrinked = GUICtrlCreateCheckbox("Dock & Shrink Bot on Start", 8, 175)
	GUICtrlSetTip(-1, "Auto Dock and Shrink the Bot Window to the Android when the Bot launches")
	$hChk_DpiAwarness = GUICtrlCreateCheckbox("DPI Awareness", 160, 135)
	GUICtrlSetTip(-1, "Launch the Bot in DPI Awareness Mode")
	$hChk_DebugMode = GUICtrlCreateCheckbox("Debug Mode", 160, 155)
	GUICtrlSetTip(-1, "Launch the Bot in Debug Mode")
	$hChk_MiniGUIMode = GUICtrlCreateCheckbox("Mini GUI Mode", 160, 175)
	GUICtrlSetTip(-1, "Launch the Bot in Mini GUI Mode")
	$hChk_HideAndroid = GUICtrlCreateCheckbox("Hide Android", 8, 195)
	GUICtrlSetTip(-1, "Hide the Android Emulator Window")

	If StringLen($g_sIniParameters) < $g_iParameters Then
		Local $iCount = $g_iParameters - StringLen($g_sIniParameters)
		For $i = 0 To $iCount
			$g_sIniParameters &= 0
		Next
	EndIf

	Local $aParameters = StringSplit($g_sIniParameters, "", 2)
	If $aParameters[0] = 1 Then GUICtrlSetState($hChk_NoWatchdog, $GUI_CHECKED)
	If $aParameters[1] = 1 Then GUICtrlSetState($hChk_StartBotDocked, $GUI_CHECKED)
	If $aParameters[2] = 1 Then GUICtrlSetState($hChk_StartBotDockedAndShrinked, $GUI_CHECKED)
	If $aParameters[3] = 1 Then GUICtrlSetState($hChk_DpiAwarness, $GUI_CHECKED)
	If $aParameters[4] = 1 Then GUICtrlSetState($hChk_DebugMode, $GUI_CHECKED)
	If $aParameters[5] = 1 Then GUICtrlSetState($hChk_MiniGUIMode, $GUI_CHECKED)
	If $aParameters[6] = 1 Then GUICtrlSetState($hChk_HideAndroid, $GUI_CHECKED)

	$hBtn_Save = GUICtrlCreateButton("Save and Close", 76, 230, 97, 25, $WS_GROUP)
	GUISetState()

	Switch $g_sIniEmulator
		Case "BlueStacks5"
			GUICtrlSetData($hCmb_Emulator, "MEmu|Nox")
		Case "MEmu"
			GUICtrlSetData($hCmb_Emulator, "BlueStacks5|Nox")
		Case "Nox"
			GUICtrlSetData($hCmb_Emulator, "BlueStacks5|MEmu")
		Case Else
			MsgBox($MB_OK, "Error", "Oops, as it looks like you changed Data in the Config File.Pleae delete all corrupted Sections!", 0, $g_hGUI_Edit)
	EndSwitch

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				GUIDelete($g_hGUI_Edit)
				ExitLoop
			Case $hCmb_Emulator
				$sSelectedEmulator = GUICtrlRead($hCmb_Emulator)
				If $sSelectedEmulator = "BlueStacks" Then
					GUICtrlSetState($hIpt_Instance, $GUI_DISABLE)
					GUICtrlSetData($hIpt_Instance, "")
				ElseIf $sSelectedEmulator <> "BlueStacks" And "BlueStacks5" Then
					GUICtrlSetState($hIpt_Instance, $GUI_ENABLE)
					Switch $sSelectedEmulator
						Case "BlueStacks5"
							GUICtrlSetData($hIpt_Instance, "Pie64_")
						Case "MEmu"
							GUICtrlSetData($hIpt_Instance, "MEmu_")
						Case "Nox"
							GUICtrlSetData($hIpt_Instance, "Nox_")
						Case Else
							MsgBox($MB_OK, "Error", "Oops, as it looks like you changed Data in the Config File. Please revert it or delete all corrupted Sections!", 0, $g_hGUI_Edit)
					EndSwitch
				EndIf
			Case $hBtn_Folder
				Local $sSelectedFolder = FileSelectFolder("Select your MultiBot Folder", $g_sIniDir)
			Case $hBtn_Save
				$sSelectedProfile = GUICtrlRead($hIpt_Profile)
				$sSelectedEmulator = GUICtrlRead($hCmb_Emulator)
				$sSelectedInstance = GUICtrlRead($hIpt_Instance)
				$iEndResult &= GUICtrlRead($hChk_NoWatchdog) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_StartBotDocked) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_StartBotDockedAndShrinked) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_DpiAwarness) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_DebugMode) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_MiniGUIMode) = $GUI_CHECKED ? 1 : 0
				$iEndResult &= GUICtrlRead($hChk_HideAndroid) = $GUI_CHECKED ? 1 : 0
				IniDelete($g_sDirProfiles, $sLstbx_SelItem)
				IniWrite($g_sDirProfiles, $sSelectedProfile, "Profile", $sSelectedProfile)
				IniWrite($g_sDirProfiles, $sSelectedProfile, "Emulator", $sSelectedEmulator)
				IniWrite($g_sDirProfiles, $sSelectedProfile, "Instance", $sSelectedInstance)
				IniWrite($g_sDirProfiles, $sSelectedProfile, "Dir", $sSelectedFolder)
				IniWrite($g_sDirProfiles, $sSelectedProfile, "Parameters", $iEndResult)
				GUIDelete($g_hGUI_Edit)
				ExitLoop
		EndSwitch
	WEnd
EndFunc   ;==>GUI_Edit


Func GUI_AutoStart()
	Local $sText = ""

	$g_hGUI_AutoStart = GUICreate("Auto Start Setup", 258, 187, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
	$g_hLst_AutoStart = GUICtrlCreateList("", 8, 48, 241, 84, BitOR($LBS_SORT, $LBS_NOTIFY, $LBS_NOSEL))
	GUICtrlCreateLabel("Add or Remove a Profile from Autostart" & @CRLF & "These are currently in the Folder:", 8, 8, 220, 34)
	$hBtn_Add = GUICtrlCreateButton("Add", 8, 144, 97, 25, $WS_GROUP)
	$hBtn_Remove = GUICtrlCreateButton("Remove", 152, 144, 97, 25, $WS_GROUP)
	GUISetState()

	$Lstbx_Sel = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
	If $Lstbx_Sel[0] > 0 Then
		For $i = 1 To $Lstbx_Sel[0]
			$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $Lstbx_Sel[$i])
			If $sLstbx_SelItem <> "" Then
				ReadIni($sLstbx_SelItem)
				If FileExists(@StartupDir & "\MultiBot -" & $g_sIniProfile & ".lnk") = 0 Then
					GUICtrlSetState($hBtn_Add, $GUI_ENABLE)
					GUICtrlSetState($hBtn_Remove, $GUI_DISABLE)
				Else
					GUICtrlSetState($hBtn_Add, $GUI_DISABLE)
					GUICtrlSetState($hBtn_Remove, $GUI_ENABLE)
				EndIf
			EndIf
		Next
	EndIf

	UpdateList_AS()

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				GUIDelete($g_hGUI_AutoStart)
				ExitLoop
			Case $hBtn_Add
				$Lstbx_Sel = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
				If $Lstbx_Sel[0] > 0 Then
					For $i = 1 To $Lstbx_Sel[0]
						$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $Lstbx_Sel[$i])
						If $sLstbx_SelItem <> "" Then
							ReadIni($sLstbx_SelItem)
							If FileExists($g_sIniDir & "\" & $g_sBotFile) = 1 Then
								$g_sBotFile = $g_sBotFile
							ElseIf FileExists($g_sIniDir & "\" & $g_sBotFileAU3) = 1 Then
								$g_sBotFile = $g_sBotFileAU3
							EndIf

							FileCreateShortcut($g_sIniDir & "\" & $g_sBotFile, @StartupDir & "\MultiBot -" & $g_sIniProfile & ".lnk", $g_sIniDir, $g_sIniProfile & " " & $g_sIniEmulator & " " & $g_sIniInstance, "Shortcut for Bot Profile:" & $g_sIniProfile)

							UpdateList_AS()
							If FileExists(@StartupDir & "\MultiBot -" & $g_sIniProfile & ".lnk") = 0 Then
								GUICtrlSetState($hBtn_Remove, $GUI_DISABLE)
								GUICtrlSetState($hBtn_Add, $GUI_ENABLE)
							Else
								GUICtrlSetState($hBtn_Remove, $GUI_ENABLE)
								GUICtrlSetState($hBtn_Add, $GUI_DISABLE)
							EndIf
						EndIf
					Next
				EndIf
			Case $hBtn_Remove
				$Lstbx_Sel = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
				If $Lstbx_Sel[0] > 0 Then
					For $i = 1 To $Lstbx_Sel[0]
						$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $Lstbx_Sel[$i])
						If $sLstbx_SelItem <> "" Then
							ReadIni($sLstbx_SelItem)
							If FileExists(@StartupDir & "\MultiBot -" & $g_sIniProfile & ".lnk") = 1 Then
								FileDelete(@StartupDir & "\MultiBot -" & $g_sIniProfile & ".lnk")
								GUICtrlSetData($g_hLst_AutoStart, "")
							EndIf
							UpdateList_AS()
							If FileExists(@StartupDir & "\MultiBot -" & $g_sIniProfile & ".lnk") = 0 Then
								GUICtrlSetState($hBtn_Remove, $GUI_DISABLE)
								GUICtrlSetState($hBtn_Add, $GUI_ENABLE)
							Else
								GUICtrlSetState($hBtn_Remove, $GUI_ENABLE)
								GUICtrlSetState($hBtn_Add, $GUI_DISABLE)
							EndIf
						EndIf

					Next
				EndIf
		EndSwitch
	WEnd
EndFunc   ;==>GUI_AutoStart

Func RunSetup()
	$Lstbx_Sel = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
	If $Lstbx_Sel[0] > 0 Then
		For $i = 1 To $Lstbx_Sel[0]
			$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $Lstbx_Sel[$i])
			If $sLstbx_SelItem <> "" Then
				ReadIni($sLstbx_SelItem)
				Local $sEmulator = $g_sIniEmulator
				If $g_sIniEmulator = "BlueStacks3" Then $sEmulator = "BlueStacks5"
				$aParameters = StringSplit($g_sIniParameters, "")
				Local $sSpecialParameter = $aParameters[1] = 1 ? " /nowatchdog" : "" & $aParameters[2] = 1 ? " /dock1" : "" & $aParameters[3] = 1 ? " /dock2" : "" & $aParameters[4] = 1 ? " /dpiaware" : "" & $aParameters[5] = 1 ? " /debug" : "" & $aParameters[6] = 1 ? " /minigui" : "" & $aParameters[7] = 1 ? " /hideandroid" : ""
				_GUICtrlStatusBar_SetText($g_hLog, "Running: " & $g_sIniProfile)
				If FileExists($g_sIniDir & "\" & $g_sBotFile) = 1 Then
					ShellExecute($g_sBotFile, $g_sIniProfile & " " & $sEmulator & " " & $g_sIniInstance & $sSpecialParameter, $g_sIniDir)
				ElseIf FileExists($g_sIniDir & "\" & $g_sBotFileAU3) = 1 Then
					ShellExecute($g_sBotFileAU3, $g_sIniProfile & " " & $sEmulator & " " & $g_sIniInstance & $sSpecialParameter, $g_sIniDir)
				Else
					MsgBox($MB_OK, "No Bot found", "Couldn't find any Bot in the Directory, please check if you have the multibot.run.exe or the multibot.run.au3 in the Dir and if you selected the right Dir!", 0, $g_hGui_Main)
					_GUICtrlStatusBar_SetText($g_hLog, "Error while Running")
				EndIf
			EndIf
		Next
	EndIf
EndFunc   ;==>RunSetup

Func CreateShortcut()
	Local $iCreatedSC = 0, $sBotFileName, $hSC
	$Lstbx_Sel = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
	ConsoleWrite(">>>>>>>>>>> $Lstbx_Sel: " & _ArrayToString($Lstbx_Sel) & @CRLF)
	If $Lstbx_Sel[0] > 0 Then
		For $i = 1 To $Lstbx_Sel[0]
			$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $Lstbx_Sel[$i])
			If $sLstbx_SelItem <> "" Then
				ReadIni($sLstbx_SelItem)
				Local $sEmulator = $g_sIniEmulator
				$aParameters = StringSplit($g_sIniParameters, "")
				Local $sSpecialParameter = $aParameters[1] = 1 ? " /nowatchdog" : "" & $aParameters[2] = 1 ? " /dock1" : "" & $aParameters[3] = 1 ? " /dock2" : "" & $aParameters[4] = 1 ? " /dpiaware" : "" & $aParameters[5] = 1 ? " /debug" : "" & $aParameters[6] = 1 ? " /minigui" : "" & $aParameters[7] = 1 ? " /hideandroid" : ""
				If FileExists($g_sIniDir & "\" & $g_sBotFile) Then
					$sBotFileName = $g_sBotFile
				ElseIf FileExists($g_sIniDir & "\" & $g_sBotFileAU3) Then
					$sBotFileName = $g_sBotFileAU3
				Else
					MsgBox($MB_OK, "No Bot found", "Couldn't find any Bot in the Directory, please check if you have the multibot.run.exe or the multibot.run.au3 in the Dir and if you selected the right Dir!", 0, $g_hGui_Main)
				EndIf
				ConsoleWrite(">>>>>>>>>>> file: " & $g_sIniDir & "\" & $sBotFileName & @CRLF)
				Local $sProfile = StringSplit($g_sIniProfile, "\")
				ConsoleWrite(">>>>>>>>>>> lnk: " & @DesktopDir & "\" & $sProfile[UBound($sProfile) - 1] & ".lnk" & @CRLF)
				$hSC = FileCreateShortcut($g_sIniDir & "\" & $sBotFileName, @DesktopDir & "\" & $sProfile[UBound($sProfile) - 1] & ".lnk", $g_sIniDir, $sProfile[UBound($sProfile) - 1] & " " & $sEmulator & " " & $g_sIniInstance & $sSpecialParameter, "Shortcut for Bot Profile:" & $g_sIniProfile)
				If $hSC = 1 Then $iCreatedSC += 1
				ConsoleWrite(">>>>>>>>>>> $hSC: " & $hSC & @CRLF)
			EndIf
		Next
		If $iCreatedSC = 1 Then
			_GUICtrlStatusBar_SetText($g_hLog, "Created " & $iCreatedSC & " Shortcut")
		ElseIf $iCreatedSC > 1 Then
			_GUICtrlStatusBar_SetText($g_hLog, "Created " & $iCreatedSC & " Shortcuts")
		EndIf
		$iCreatedSC = 0
	EndIf
EndFunc   ;==>CreateShortcut

Func UpdateList_Main() ; Main List Updating
	Local $j = 0, $aSections

	GetBotVers()

	_GUICtrlListView_BeginUpdate($g_hListview_Main)
	_GUICtrlListView_DeleteAllItems($g_hListview_Main)
	$aSections = IniReadSectionNames($g_sDirProfiles)
	For $i = 1 To UBound($aSections, 1) - 1
		If $aSections[$i] <> "Options" Then
			_GUICtrlListView_AddItem($g_hListview_Main, $aSections[$i])
			$iSectionVers = IniRead($g_sDirProfiles, $aSections[$i], "BotVers", "")
			_GUICtrlListView_AddSubItem($g_hListview_Main, $j, $iSectionVers, 1)
			$j += 1
		EndIf
	Next
	_GUICtrlListView_EndUpdate($g_hListview_Main)
EndFunc   ;==>UpdateList_Main



Func UpdateList_AS()
	Local $aSections

	GUICtrlSetData($g_hLst_AutoStart, "")
	$aSections = IniReadSectionNames($g_sDirProfiles)
	If @error <> 0 Then
	Else
		For $i = 1 To $aSections[0]
			$sProfiles = IniRead($g_sDirProfiles, $aSections[$i], "Profile", "")
			If FileExists(@StartupDir & "\MultiBot -" & $sProfiles & ".lnk") Then
				GUICtrlSetData($g_hLst_AutoStart, $sProfiles)
			EndIf
		Next
	EndIf

EndFunc   ;==>UpdateList_AS

Func GetBotVers()
	Local $aSections

	$aSections = IniReadSectionNames($g_sDirProfiles)
	For $i = 1 To UBound($aSections, 1) - 1
		ReadIni($aSections[$i])
		$hBotVers = FileOpen($g_sIniDir & "\multibot.run.au3")

		$sBotVers = FileRead($hBotVers)
		$aBotVers = StringRegExp($sBotVers, 'v[0-9](.*")', 2)

		FileClose($hBotVers)
		If $aSections[$i] <> "Options" Then
			If IsArray($aBotVers) Then
				IniWrite($g_sDirProfiles, $aSections[$i], "BotVers", StringReplace($aBotVers[0], '"', ""))
			Else
				$hBotVers = FileOpen($g_sIniDir & "\multibot.run.version.au3")
				$sBotVers = FileRead($hBotVers)
				$aBotVers = StringRegExp($sBotVers, 'v[0-9](.*")', 2)
				FileClose($hBotVers)
				If IsArray($aBotVers) Then IniWrite($g_sDirProfiles, $aSections[$i], "BotVers", StringReplace($aBotVers[0], '"', ""))
			EndIf
		EndIf
	Next
EndFunc   ;==>GetBotVers

Func ReadIni($sSelectedProfile)

	$g_sIniProfile = IniRead($g_sDirProfiles, $sSelectedProfile, "Profile", "")
	If StringRegExp($g_sIniProfile, " ") = 1 Then
		$g_sIniProfile = '"' & $g_sIniProfile & '"'
	EndIf
	$g_sIniEmulator = IniRead($g_sDirProfiles, $sSelectedProfile, "Emulator", "")
	$g_sIniInstance = IniRead($g_sDirProfiles, $sSelectedProfile, "Instance", "")
	$g_sIniDir = IniRead($g_sDirProfiles, $sSelectedProfile, "Dir", "")

	Local $iParam
	For $i = 0 To $g_iParameters
		$iParam &= 0
	Next
	$g_sIniParameters = IniRead($g_sDirProfiles, $sSelectedProfile, "Parameters", $iParam)

EndFunc   ;==>ReadIni

Func WM_CONTEXTMENU($hWnd, $msg, $wParam, $lParam)
	Local $tPoint = _WinAPI_GetMousePos(True, GUICtrlGetHandle($g_hListview_Main))
	Local $iY = DllStructGetData($tPoint, "Y")

	$lst2 = _GUICtrlListView_GetItemCount($g_hListview_Main)
	If $lst2 > 0 Then
		For $i = 0 To 50
			$iLstbx_GetSel = _GUICtrlListView_GetSelectedCount($g_hListview_Main)
			If $iLstbx_GetSel > 1 Then
				_GUICtrlMenu_SetItemDisabled($g_hContext_Main, 1)
				_GUICtrlMenu_SetItemDisabled($g_hContext_Main, 3)
			ElseIf $iLstbx_GetSel < 2 Then
				_GUICtrlMenu_SetItemEnabled($g_hContext_Main, 1)
				_GUICtrlMenu_SetItemEnabled($g_hContext_Main, 3)
			EndIf

			Local $aRect = _GUICtrlListView_GetItemRect($g_hListview_Main, $i)
			If ($iY >= $aRect[1]) And ($iY <= $aRect[3]) Then _ContextMenu($i)
		Next
		Return $GUI_RUNDEFMSG

	ElseIf _GUICtrlListView_GetItemCount($g_hListview_Main) < 1 Then
		Return
	EndIf
EndFunc   ;==>WM_CONTEXTMENU

Func _ContextMenu($sItem)
	Switch _GUICtrlMenu_TrackPopupMenu($g_hContext_Main, GUICtrlGetHandle($g_hListview_Main), -1, -1, 1, 1, 2)
		Case $eRun
			RunSetup()

		Case $eEdit
			GUISetState(@SW_DISABLE, $g_hGui_Main)
			$g_aGuiPos_Main = WinGetPos($g_hGui_Main)
			_GUICtrlStatusBar_SetText($g_hLog, "Begin to edit Setup")
			GUI_Edit()
			UpdateList_Main()
			_GUICtrlStatusBar_SetText($g_hLog, "Setup Edit done")
			GUISetState(@SW_ENABLE, $g_hGui_Main)

		Case $eDelete

			$Lstbx_Sel = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
			If $Lstbx_Sel[0] > 0 Then
				For $i = 1 To $Lstbx_Sel[0]
					$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $Lstbx_Sel[$i])
					If $sLstbx_SelItem <> "" Then
						IniDelete(@MyDocumentsDir & "\Profiles.ini", $sLstbx_SelItem)
						If FileExists(@StartupDir & "\MultiBot -" & $sLstbx_SelItem & ".lnk") = 1 Then FileDelete(@StartupDir & "\MultiBot -" & $sLstbx_SelItem & ".lnk")
					EndIf
				Next
			EndIf
			If $Lstbx_Sel[0] = 1 Then
				_GUICtrlStatusBar_SetText($g_hLog, "Deleted " & $Lstbx_Sel[0] & " Setup")
			Else
				_GUICtrlStatusBar_SetText($g_hLog, "Deleted " & $Lstbx_Sel[0] & " Setups")
			EndIf
			UpdateList_Main()

		Case $eNickname

			$sLstbx_GetSelTxt = _GUICtrlListView_GetItemTextArray($g_hListview_Main)
			ReadIni($sLstbx_GetSelTxt[1])
			$sIptbx = InputBox("Give Profile a Nickname", "Here you can set a NickName for each Setup. It will be shown in Brackets behind the Profile Name! To Remove a Nick just press OK when nothing is in the Input!", "", "", -1, -1, Default, Default, 0, $g_hGui_Main)
			If $sIptbx = "" Then
				IniRenameSection($g_sDirProfiles, $sLstbx_GetSelTxt[1], $g_sIniProfile)
			Else
				IniRenameSection($g_sDirProfiles, $sLstbx_GetSelTxt[1], $g_sIniProfile & " ( " & $sIptbx & " )")
			EndIf
			_GUICtrlStatusBar_SetText($g_hLog, "Setup renamed!")
			UpdateList_Main()

	EndSwitch
EndFunc   ;==>_ContextMenu

Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)

	Local $hWndFrom, $iCode, $tNMHDR, $hWndListView
	$hWndListView = $g_hListview_Main
	If Not IsHWnd($g_hListview_Main) Then $hWndListView = GUICtrlGetHandle($g_hListview_Main)

	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iCode = DllStructGetData($tNMHDR, "Code")
	Switch $hWndFrom
		Case $hWndListView
			Switch $iCode

				Case $NM_DBLCLK
					Local $tInfo = DllStructCreate($tagNMITEMACTIVATE, $ilParam)

					$Index = DllStructGetData($tInfo, "Index")

					$subitemNR = DllStructGetData($tInfo, "SubItem")
					If $Index <> -1 Then
						If _IsPressed(10) Then
							$Lstbx_Sel = _GUICtrlListView_GetItemText($g_hListview_Main, $Index)
							ReadIni($Lstbx_Sel)
							ToolTip("Profile: " & $g_sIniProfile & @CRLF & "Emulator: " & $g_sIniEmulator & @CRLF & "Instance: " & $g_sIniInstance & @CRLF & "Directory: " & $g_sIniDir)
							Do
								Sleep(100)
							Until _IsPressed(10) = False
							ToolTip("")
						Else
							RunSetup()
						EndIf
					EndIf

			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

Func UpdateSelect()

	FileMove(@ScriptDir & "\" & @ScriptName, @ScriptDir & "\" & "SelectMultiBotRunOLD" & $g_sVersion & ".exe")
	$hUpdateFile = InetGet("https://github.com/promac2k/SelectMultiBotRun/raw/master/SelectMultiBotRun.Exe", @ScriptDir & "\SelectMultiBotRun.exe", $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	Do
		Sleep(250)
	Until InetGetInfo($hUpdateFile, $INET_DOWNLOADCOMPLETE)

	InetClose($hUpdateFile)
	MsgBox($MB_OK, "Update", "Update finished! Downloaded newer Version and placed it in old Directoy!.New Version gets now started! Just delete the SelectMultiBotRunOLD" & $g_sVersion & ".exe and continue botting :)", 0, $g_hGui_Main)
	ShellExecute(@ScriptDir & "\SelectMultiBotRun.exe")
	Exit

EndFunc   ;==>UpdateSelect

Func CheckUpdate()
	$sTempPath = @MyDocumentsDir & "\SelectMultiBotRun_Info.txt"
	$hUpdateFile = InetGet("https://raw.githubusercontent.com/tehbank/SelectMultiBotRun/master/SelectMultiBotRun_Info.txt", $sTempPath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	Do
		Sleep(250)
	Until InetGetInfo($hUpdateFile, $INET_DOWNLOADCOMPLETE)

	InetClose($hUpdateFile)
	$hGitVersion = IniRead($sTempPath, "General", "DisplayVers", "")
	$sGitVersion = StringStripWS($hGitVersion, 8)
	$Update = _VersionCompare($g_sVersion, $sGitVersion)

	If $Update = -1 Then GUICtrlSetState($g_hLblUpdateAvailable, $GUI_SHOW)

	FileDelete($sTempPath)
EndFunc   ;==>CheckUpdate

Func ChangeLog()
	Local $sTitle, $sMessage, $sDate

	$sTempPath = @MyDocumentsDir & "SelectMultiBotRun_Info.txt"
	$hUpdateFile = InetGet("https://raw.githubusercontent.com/promac2k/SelectMultiBotRun/master/SelectMultiBotRun_Info.txt", $sTempPath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	Do
		Sleep(250)
	Until InetGetInfo($hUpdateFile, $INET_DOWNLOADCOMPLETE)

	InetClose($hUpdateFile)
	If Not FileExists($g_sDirProfiles) Then Return
	$hGitVersion = IniRead($sTempPath, "General", "DisplayVers", "")
	$sDisplayVers = StringStripWS($hGitVersion, 8)
	$sDisplayVersSent = IniRead($g_sDirProfiles, "Options", "DisplayVersSent", "")
	If _VersionCompare($sDisplayVers, $g_sVersion) = 0 And _VersionCompare($sDisplayVersSent, $g_sVersion) = -1 Then
		$sTitle = IniRead($sTempPath, "Changelog", "Title", "")
		$sMessage = StringReplace(IniRead($sTempPath, "Changelog", "Message", ""), "%", @CRLF)
		$sDate = IniRead($sTempPath, "Changelog", "Date", "")
		IniWrite($g_sDirProfiles, "Options", "DisplayVersSent", $g_sVersion)
		GUISetState(@SW_DISABLE, $g_hGui_Main)
		GUI_ChangeLog($sTitle, $sMessage, $sDate)
		GUISetState(@SW_ENABLE, $g_hGui_Main)
	EndIf

	FileDelete($sTempPath)
EndFunc   ;==>ChangeLog

Func GUI_ChangeLog($Title, $Message, $Date)

	$g_aGuiPos_Main = WinGetPos($g_hGui_Main)

	$g_hGUI_ChangeLog = GUICreate($Title, 258, 230, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
	GUICtrlCreateLabel($Date, 8, 8)
	GUICtrlCreateEdit($Message, 8, 30, 242, 160, $ES_READONLY)
	GUICtrlSetBkColor(-1, $COLOR_WHITE)
	$hBtn_Dismiss = GUICtrlCreateButton("Dismiss", 76, 200, 97, 25, $WS_GROUP)
	GUISetState()

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $hBtn_Dismiss
				GUIDelete($g_hGUI_ChangeLog)
				ExitLoop
		EndSwitch
	WEnd

EndFunc   ;==>GUI_ChangeLog

; THANKS COSOTE
#Region Android

Func GetMEmuPath()
	Local $sMEmuPath = EnvGet("MEmu_Path") & "\MEmu\"
	If FileExists($sMEmuPath & "MEmu.exe") = 0 Then
		Local $sInstallLocation = RegRead($HKLM & "\SOFTWARE" & $Wow6432Node & "\Microsoft\Windows\CurrentVersion\Uninstall\MEmu\", "InstallLocation")
		If @error = 0 And FileExists($sInstallLocation & "\MEmu\MEmu.exe") = 1 Then
			$sMEmuPath = $sInstallLocation & "\MEmu\"
		Else
			Local $sDisplayIcon = RegRead($HKLM & "\SOFTWARE" & $Wow6432Node & "\Microsoft\Windows\CurrentVersion\Uninstall\MEmu\", "DisplayIcon")
			If @error = 0 Then
				Local $iLastBS = StringInStr($sDisplayIcon, "\", 0, -1)
				$sMEmuPath = StringLeft($sDisplayIcon, $iLastBS)
				If StringLeft($sMEmuPath, 1) = """" Then $sMEmuPath = StringMid($sMEmuPath, 2)
			Else
				$sMEmuPath = @ProgramFilesDir & "\Microvirt\MEmu\"
			EndIf
		EndIf
	EndIf
	Return $sMEmuPath
EndFunc   ;==>GetMEmuPath

Func GetNoxPath()
	Local $sNoxPath = RegRead($HKLM & "\SOFTWARE" & $Wow6432Node & "\DuoDianOnline\SetupInfo\", "InstallPath")
	If @error = 0 Then
		If StringRight($sNoxPath, 1) <> "\" Then $sNoxPath &= "\"
		$sNoxPath &= "bin\"
	Else
		$sNoxPath = ""
	EndIf
	Return $sNoxPath
EndFunc   ;==>GetNoxPath

Func GetNoxRtPath()
	Local $sNoxRtPath = RegRead($HKLM & "\SOFTWARE\BigNox\VirtualBox\", "InstallDir")
	If @error = 0 Then
		If StringRight($sNoxRtPath, 1) <> "\" Then $sNoxRtPath &= "\"
	Else
		$sNoxRtPath = @ProgramFilesDir & "\Bignox\BigNoxVM\RT\"
	EndIf
	Return $sNoxRtPath
EndFunc   ;==>GetNoxRtPath

Func GetBlueStacksPath()
	$sBlueStacksPath = RegRead($HKLM & "\SOFTWARE\BlueStacks_nxt", "InstallDir")
	$sPlusMode = RegRead($HKLM & "\SOFTWARE\BlueStacks_nxt\", "Engine") = "plus"
	$sFrontend = "HD-Frontend.exe"
	If $sPlusMode Then $sFrontend = "HD-Plus-Frontend.exe"
	If $sBlueStacksPath = "" And FileExists(@ProgramFilesDir & "\BlueStacks_nxt\" & $sFrontend) = 1 Then
		$sBlueStacksPath = @ProgramFilesDir & "\BlueStacks_nxt\"
	EndIf

	Return $sBlueStacksPath
EndFunc   ;==>GetBlueStacksPath

Func GetiToolsPath()
	DebugLog("GetiToolsPath")
	Local $siTools_Path = ""
	If $siTools_Path <> "" And FileExists($siTools_Path & "\iToolsAVM.exe") = 0 Then
		$siTools_Path = ""
	EndIf
	Local $sInstallLocation = ""
	Local $sDisplayIcon = RegRead($HKLM & "\SOFTWARE" & $Wow6432Node & "\Microsoft\Windows\CurrentVersion\Uninstall\iToolsAVM\", "DisplayIcon")
	DebugLog("$sDisplayIcon " & $sDisplayIcon)
	If @error = 0 Then
		Local $iLastBS = StringInStr($sDisplayIcon, "\", 0, -1) - 1
		$sInstallLocation = StringLeft($sDisplayIcon, $iLastBS)
	EndIf
	If $siTools_Path = "" And FileExists($sInstallLocation & "\iToolsAVM.exe") = 1 Then
		$siTools_Path = $sInstallLocation
	EndIf
	If $siTools_Path = "" And FileExists(@ProgramFilesDir & "\iToolsAVM\iToolsAVM.exe") = 1 Then
		$siTools_Path = @ProgramFilesDir & "\iToolsAVM"
	EndIf
	SetError(0, 0, 0)
	If $siTools_Path <> "" And StringRight($siTools_Path, 1) <> "\" Then $siTools_Path &= "\"
	Return StringReplace($siTools_Path, "\\", "\")
EndFunc   ;==>GetiToolsPath

Func IsAndroidInstalled($sAndroid)
	Local $sPath, $sFile, $bIsInstalled = False
	DebugLog("Android to Check -> " & $sAndroid)

	Switch $sAndroid
		Case "MEmu"
			DebugLog("MEmu")
			$sPath = GetMEmuPath()
			$sFile = "MEmu.exe"
		Case "Nox"
			DebugLog("Nox")
			$sPath = GetNoxPath()
			$sFile = "Nox.exe"
		Case "BlueStacks5"
			DebugLog("BlueStacks")
			$sPath = GetBlueStacksPath()
			$bPlusMode = RegRead($HKLM & "\SOFTWARE\BlueStacks_nxt\", "Engine") = "plus"
			$sFile = "HD-Frontend.exe"
			If $bPlusMode Then $sFile = "HD-Plus-Frontend.exe"
		Case "iTools"
			DebugLog("iTools")
			$sPath = GetiToolsPath()
			$sFile = "iToolsAVM.exe"
	EndSwitch

	If FileExists($sPath & $sFile) = 1 Then $bIsInstalled = True
	DebugLog($sAndroid & " Is Installed ?" & $bIsInstalled)
	Return $bIsInstalled
EndFunc   ;==>IsAndroidInstalled

Func GetInstanceMgrPath($sAndroid)

	Local $sManagerPath

	Switch $sAndroid
		Case "BlueStacks5"
			$sManagerPath = GetBlueStacksPath() & "BstkVMMgr.exe"
		Case "MEmu"
			$sManagerPath = EnvGet("MEmuHyperv_Path") & "\MEmuManage.exe"
			If FileExists($sManagerPath) = 0 Then
				$sManagerPath = GetMEmuPath() & "..\MEmuHyperv\MEmuManage.exe"
			EndIf
		Case "Nox"
			$sManagerPath = GetNoxRtPath() & "BigNoxVMMgr.exe"
		Case "iTools"
			$sVirtualBox_Path = RegRead($HKLM & "\SOFTWARE\Oracle\VirtualBox\", "InstallDir")
			If @error <> 0 And FileExists(@ProgramFilesDir & "\Oracle\VirtualBox\") Then
				$sVirtualBox_Path = @ProgramFilesDir & "\Oracle\VirtualBox\"
			EndIf
			$sVirtualBox_Path = StringReplace($sVirtualBox_Path, "\\", "\")
			$sManagerPath = $sVirtualBox_Path & "VBoxManage.exe"
	EndSwitch
	Return $sManagerPath
EndFunc   ;==>GetInstanceMgrPath

Func DebugLog($Text)
	ConsoleWrite(TimeDebug() & " - " & $Text & @CRLF)
EndFunc   ;==>DebugLog

Func TimeDebug() ;Gives the time in '[14:00:00.000]' format
	Return "[" & @YEAR & "-" & @MON & "-" & @MDAY & " " & _NowTime(5) & "." & @MSEC & "] "
EndFunc   ;==>TimeDebug
#EndRegion Android

#Region CMD
Func LaunchConsole($sCMD, $sParameter, $bProcessKilled, $iTimeOut = 10000)
	Local $sData, $iPID, $hTimer
	If StringLen($sParameter) > 0 Then $sCMD &= " " & $sParameter
	$hTimer = TimerInit()
	$bProcessKilled = False
	$iPID = Run($sCMD, "", @SW_HIDE, $STDERR_MERGED)
	If $iPID = 0 Then
		Return
	EndIf
	Local $hProcess
	If _WinAPI_GetVersion() >= 6.0 Then
		$hProcess = _WinAPI_OpenProcess($PROCESS_QUERY_LIMITED_INFORMATION, 0, $iPID)
	Else
		$hProcess = _WinAPI_OpenProcess($PROCESS_QUERY_INFORMATION, 0, $iPID)
	EndIf

	$sData = ""
	$iDelaySleep = 100
	Local $iTimeOut_Sec = Round($iTimeOut / 1000)

	While True
		If $hProcess Then
			_WinAPI_WaitForSingleObject($hProcess, $iDelaySleep)
		Else
			Sleep($iDelaySleep)
		EndIf
		$sData &= StdoutRead($iPID)
		If @error Then ExitLoop
		If ($iTimeOut > 0 And TimerDiff($hTimer) > $iTimeOut) Then ExitLoop
	WEnd

	If $hProcess Then
		_WinAPI_CloseHandle($hProcess)
		$hProcess = 0
	EndIf

	CleanLaunchOutput($sData)

	If ProcessExists($iPID) Then
		If ProcessClose($iPID) = 1 Then
			$bProcessKilled = True
		EndIf
	EndIf
	StdioClose($iPID)
	Return $sData
EndFunc   ;==>LaunchConsole

Func CleanLaunchOutput(ByRef $output)

	$output = StringReplace($output, @CR & @CR, "")
	$output = StringReplace($output, @CRLF & @CRLF, "")
	If StringRight($output, 1) = @LF Then $output = StringLeft($output, StringLen($output) - 1)
	If StringRight($output, 1) = @CR Then $output = StringLeft($output, StringLen($output) - 1)
EndFunc   ;==>CleanLaunchOutput

#EndRegion CMD
