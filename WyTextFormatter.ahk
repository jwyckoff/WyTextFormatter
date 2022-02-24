;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         Jason Wyckoff <jason@jasonwyckoff.com>  
;
; 
;**************************************************************************************
; Script configuration settings
;**************************************************************************************
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force ; only 1 instance of this script can be loaded.  Always reload it.
SetTimer, R_CheckReload, 2000
;**************************************************************************************
;
; 	Global Stuff - Does not need to change
;
;**************************************************************************************

;SplitPath, A_ScriptFullPath, name, dir, ext, name_no_ext, drive
SplitPath, A_ScriptFullPath,,,, WyApp
;Menu,Tray,Togglecheck,Run At Startup
wyAppName = %WyApp% by Jason Wyckoff


;-------+
;TRAY   |
;-------+
Menu, Tray, NoStandard
Menu, Tray, Tip, %wyAppName%
Menu, Tray, Add,
Menu, Tray, Add, Run at Startup, R_StartUp
Menu, Tray, Add
Menu, Tray, Add, Quit, R_Quit
Menu, Tray, Icon, %A_ScriptDir%/%WyApp%.ico, , 1" 

IfExist,%a_startup%/%A_ScriptName%.lnk
{
	FileDelete,%a_startup%/%A_ScriptName%.lnk
	FileCreateShortcut,%A_ScriptFullPath%,%A_Startup%/%A_ScriptName%.lnk
	Menu,Tray,Check,Run At Startup
}
;**************************************************************************************
;
; FUNCTIONS
;
;**************************************************************************************

	;---------------------------------------------------------------
	; SpaceFormatting(stringToSpace)
	;---------------------------------------------------------------
	SpaceFormatting(stringToSpace)
	{
		length := StrLen(stringToSpace)
		;MsgBox %length%
		x := 1
		
		AutoTrim, Off
		
		while (x <= length)
		{
			;MsgBox x = %x%
			;letter := SubStr(stringToSpace,1,1)
			;MsgBox Letter = %letter%
			letter := SubStr(stringToSpace,x,1)
			;MsgBox Letter = %letter%
			
			newString = %newString%%letter%%A_Space%
			
			;MsgBox %newString%
			
			x := x + 1
			;MsgBox %x%
		}
		;MsgBox %newString%
		
		AutoTrim, On
		

		return newString
	}

	;---------------------------------------------------------------
	; StripFormattingOfText()
	;---------------------------------------------------------------

	StripFormattingOfText()
	{
		Clip0 = %ClipBoardAll%
		ClipBoard = %ClipBoard%     ; Convert to text
		; Send ^v                     ; For best compatibility: SendPlay
		; Sleep 1                   	; Don't change clipboard while it is pasted! (Sleep > 0)
		; ClipBoard = %Clip0%         ; Restore original ClipBoard
		; VarSetCapacity(Clip0, 0)    ; Free memory
	}

	;---------------------------------------------------------------
	; AddSpaceBeforeUpper(stringToFormat)
	;---------------------------------------------------------------
	AddSpaceBeforeUpper(stringToFormat)
	{
		StringCaseSense, On

		length := StrLen(stringToFormat)
		;MsgBox %length%
		x := 1
		
		AutoTrim, Off
		
		while (x <= length)
		{
			;StringToFormat
			letter := SubStr(stringToFormat,x,1)
			upperLetter := ToUpper(letter)

			if (upperLetter == letter && x > 1)
			{
				; add a space because the letter is UPPPER CASE
				newString = %newString%%A_Space%
			}		
		
			newString = %newString%%letter%
			
			x := x + 1
		} 
		
		AutoTrim, On
		StringCaseSense, Off

		return newString
	}
	
	;---------------------------------------------------------------
	; CreateLineOfSameLength(stringToFormat,lineChar)
	;---------------------------------------------------------------
	CreateLineOfSameLength(stringToFormat,lineChar)
	{
		StringCaseSense, On

		length := StrLen(stringToFormat)
		;MsgBox %length%
		x := 1
		
		AutoTrim, Off
		
		while (x <= length)
		{
			newString = %newString%%lineChar%
			x := x + 1
		} 
		
		AutoTrim, On
		StringCaseSense, Off

		return newString
	}
	
	;---------------------------------------------------------------
	; FormatPhoneNumber(phonenumber)
	;---------------------------------------------------------------
	FormatPhoneNumber(phonenumber)
	{
		phonenumber := RegExReplace(phonenumber, "[^0-9]", "")
		
		if (SubStr(phonenumber, 1,1) = 1) 
		{
			phonenumber := SubStr(phonenumber,2)
		}
		
		sub1 := SubStr(phonenumber,1,3)
		sub2 := SubStr(phonenumber,4,3)
		sub3 := SubStr(phonenumber,7,4)
		sub4 := SubStr(phonenumber,11)
		phonenumber = (%sub1%) %sub2%-%sub3% 
		
		if(sub4 > "")
		{
			phonenumber = %phonenumber% x%sub4%
		}
		return phonenumber
	}
		
		
	ToUpper(stringToFormat)
	{
	   ; StringUpper, OutputVar, InputVar [, T]
	   StringUpper, returnString, stringToFormat
	   ;MsgBox %returnString%
	   return returnString
	}


	InputParse(inputText)
	;InputParse - sets the due date in MLO
	{
		Send {F2}
		Send {Right}
		Send {Space} %inputText%
		Send !{Enter}
	}

	;SetReminder - Sets the reminder time in MLO
	SetReminder(spanOfTime)
	{
		Send {F2}
		Send {Right}
		Send {Space} reminder %spanOfTime%
		Send !{Enter}
	}
	LightroomStack2Photos()
	{
		Send +{Right}
		Sleep 100
		Send ^g
		Sleep 400
		Send {Right}
	}
	CreateTimeStamp(message)
	{
		Send ^{End}
		FormatTime, TimeString, , ddd M/d/yyyy h:mm:ss tt
		Send -----------------------------{ENTER}
		Send %TimeString%{ENTER}
		Send %message%
	}

	StringJoin(arr, d=";")
	{
		retString :=
		Loop % arr.MaxIndex()
		{
			thisword := arr[a_index]
			retString = %retString%%d%%thisword%
			;MsgBox, Color number %retString%
		}
		return retString
	}

	CreateTimeStampCalled()
	{
		FormatTime, TimeString, , ddd M/d/yyyy h:mm:ss tt
		Send Called at %TimeString%{ENTER}
	}

	CreateDateTimeStamp(format)
	{
		FormatTime, TimeString2, , %format%
		return %TimeString2%
	}

	SendDateTimeStamp(format)
	{ 
		var := CreateDateTimeStamp(format)
		Send %var%
	}	


;**************************************************************************************
;
; END - FUNCTIONS
;
;**************************************************************************************
	
;**************************************************************************************
;
; GUI
;
;**************************************************************************************
{
	global MyRadioGroup
	global MyEdit
	global MyPreview
	global AutoPaste

	R_WyGui_Build:
		Gui, 1:Font, S14 CRed, Raleway
		Gui, 1:Color, 0xefe6a3
		Gui, 1:Add, Text, x5 y5 w390 h30 , WyLauncher

		Gui, 1:Font, S10 CBlack, Noto Sans
		
		Gui, 1:Add, Text, section, Text:	
		Gui, 1:Add, Edit, w390 vMyEdit,  
		
		Gui, 1:Add, Text, section, Action:	
		Gui, 1:Add, Radio, vMyRadioGroup,&1 - Format Phone Number    
		Gui, 1:Add, Radio,, &2 - Format Title
		Gui, 1:Add, Radio,, &3 - Format Upper
		Gui, 1:Add, Radio,, &4 - Format Lower
		Gui, 1:Add, Radio,, &5 - Strip Formatting
		Gui, 1:Add, Radio,, &6 - S P A C E  T I T L E
		Gui, 1:Add, Radio,, &7 - AddSpaceBeforeUpper
		Gui, 1:Add, Radio,, &8 - -------------------
		Gui, 1:Add, Checkbox, vAutoPaste, Send Control-&V?
		Gui, 1:Add, Edit, w390 vMyPreview,  
		
		Gui, 1:Add, Text, ys-	
		Gui, 1:Add, Button, h25 w126  Default,OK
		Gui, 1:Add, Button, h25 w126 , &Preview   	
		Gui, 1:Add, Button, h25 w126 , &Cancel   	
		ButtonCancel:
			Gui, 1:Hide
			return
		ButtonPreview:
			Gui, 1:Submit
			tempString := FormatStringPerFormat(MyRadioGroup,MyEdit)
			GuiControl,, MyPreview, %tempString%
			Gui, 1:Show
			return

		ButtonOK:
			Gui, 1:Submit
			tempString := FormatStringPerFormat(MyRadioGroup,MyEdit)
			clipboard := tempString

			if (AutoPaste = 1)
				Send ^v
			
		EndFx:
			return
		return

	WyGui_Show()
	{ 
		Gui, 1:Show
		GuiControl,, MyEdit, %clipboard%
		GuiControl,, MyPreview
		return
	}
			
		
	

	; "MAIN"
	goto, R_WyGui_Build
	; END "MAIN"

}

FormatStringPerFormat(iOption, inputString)
{
	returnString := ""
	switch iOption
		{
			case 1: ;Format Phone Number
				returnString := FormatPhoneNumber(inputString)
			case 2: ; Format Title
				StringUpper, returnString, inputString , T
			case 3:  ; Format Upper
				StringUpper, returnString, inputString
			case 4: ; Format Lower
				StringLower, returnString, inputString
			case 5:
				returnString := StripFormattingOfText()
			case 6: 
				StringUpper, inputString, inputString
				returnString := SpaceFormatting(inputString)
			case 7: 
				returnString := AddSpaceBeforeUpper(inputString)
			case 8: 
				returnString := CreateLineOfSameLength(inputString,"-")
		}
	return returnString
}

;**************************************************************************************
;
; END - GUI
;
;**************************************************************************************

;**************************************************************************************
;
; Global (cross-applications) search & replace strings and key remappings
;
;**************************************************************************************

	!=::
		Send ^c
		WyGui_Show()
		return

		; {
		; 	#IfWinActive , ahk_class ahk_class AutoHotkeyGUI
		; 	p::
		; 		clipboard := FormatPhoneNumber(MyEdit)
		; 		Gui, 1:Hide
		; 		return
		; 	t::
		; 		StringUpper, clipboard, MyEdit , T
		; 		Gui, 1:Hide
		; 		return
		; 	u::
		; 		clipboard := StringUpper, clipboard, MyEdit
		; 		Gui, 1:Hide
		; 		return
				
		; 	#IfWinActive
		; }


	

;**************************************************************************************
; 
; END Global (cross-applications) search & replace strings and key remappings
;
;**************************************************************************************




;**************************************************************************************
;
; 	GLOBAL SUBS
;
;**************************************************************************************
R_Quit:
	ExitApp
	return

; this routine will check if the file has changed, and reloads it automatically.  If filename has changed,
; change it below.
R_CheckReload:
	FileGetAttrib, attribs, %A_ScriptFullPath%
	AHKscripts = %A_ScriptName%    
	;MsgBox %A_ScriptName%   
	Loop, Parse, AHKscripts, `,
	{  
		script = %A_ScriptDir%\%A_LoopField%
		FileGetAttrib, attribs, %script%
		IfInString, attribs, A
		{
			FileSetAttrib, -A, %script%
			SplashTextOn, , , Updated %A_LoopField%

			IfExist, %BackupDir%
			FileCopy, %script%, %BackupDir%\%A_LoopField%-%BackupName%

			Sleep, 500
			SplashTextOff
			Reload
			Return
		}
	}
	Return
	
R_StartUp:
	Menu,Tray,Togglecheck,Run At Startup

	IfExist,%a_startup%/%A_ScriptName%.lnk
	{
		FileDelete,%a_startup%/%A_ScriptName%.lnk 
	}
	else
	{
		FileCreateShortcut,%A_ScriptFullPath%,%A_Startup%/%A_ScriptName%.lnk
	}
	return

