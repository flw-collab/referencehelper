; Referencehelper

; Copyright (C) 2019 Felix L. Wenger  

    ; This program is free software: you can redistribute it and/or modify
    ; it under the terms of the GNU General Public License as published by
    ; the Free Software Foundation, either version 3 of the License, or
    ; (at your option) any later version.

    ; This program is distributed in the hope that it will be useful,
    ; but WITHOUT ANY WARRANTY; without even the implied warranty of
    ; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    ; GNU General Public License for more details.

    ; You should have received a copy of the GNU General Public License
    ; along with this program.  If not, see <https://www.gnu.org/licenses/>.

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


;#Persistent
Menu, Tray, Tip , Referencehelper:`n`nWIN+1: Copy`nWIN+2: Insert next`nWIN+3: Reset by one`n`n(C) 2019 Felix L. Wenger
Menu, Tray, Add  ; Creates a separator line.
Menu, Tray, Add, Readme, MenuHandler ; Creates a new menu item.

;Declare variables and options
vReferenceOriginal := "" 
vReferenceNew :=
vReferenceOld :=
vReferenceOriginalTrimmed :=
LeadingZeros :=
vEnd := "" ; the current reference end
vEndNew := "" ; the new reference end
vModus := 0 ; 0 = import new reference; 1 = reset, i.e. go back one increment; 2 = output new reference to cursor selection
vCase := 0 ; the modus that this script works with, i.e., do we work with a digit or non-digit reference end
vDebug := 0 ; 1 = show debug menu

!esc::ExitApp ; Alt+Esc to close/exit script 

#1::
{ ;Win+1 to copy reference

	vCase := ImportReference_Function()
	If vDebug = 1 
	{
	MsgBox, 64, Debug, 
	( LTrim
	ImportReference_Function: Debug
	vModus: %vModus%
	vCase: %vCase%
	vReferenceOriginal: %vReferenceOriginal%
	vEnd: %vEnd%
	vReferenceOriginalTrimmed: %vReferenceOriginalTrimmed%
	)
	}
	return
}

#2::
{ ;Win+2 to insert new reference

	If vCase = 0 
	{
	MsgBox, 16, , Please use hotkey Win+1 to import a reference from a selected text first.
	Exit
	} 
	Else if (vCase = "CaseDigit") 
	{
	vModus := 2
	CaseDigit_Function()
	} 
	Else if (vCase = "CaseNonDigit") 
	{
	vModus := 2
	CaseNonDigit_Function()
	}
	return
}
	
#3::
{ ;Win+3 to reset next reference by one increment

	If vCase = 0 
	{
	MsgBox, 16, , Please use hotkey Win+1 to import a reference from a selected text first.
	Exit
	} Else { 
	vModus := 1
	CaseReset_Function()
	}

	
	If vDebug = 1 
	{
	MsgBox, 64, Debug, 
	( LTrim
	CaseReset_Function: Debug
	vModus: %vModus%
	vCase: %vCase%
	vEnd: %vEnd%
	vEndNew: %vEndNew%
	vReferenceNew: %vReferenceNew%
	)
	}
	return	
}

ImportReference_Function(){
	
	global
	
		; If vDebug = 1 
		; {
		; MsgBox, 64, Debug, 
		; ( LTrim
		; ImportReference_Function started.
		; )
		; }
	
	;##########Step 1 - Import reference from selection through the clipboard
		ClipSaved := ClipboardAll
		Clipboard := ""
		Send, ^c
		ClipWait, 1
		if (!ErrorLevel){
			AltClipboard2 = %Clipboard%
			vReferenceOriginal := AltClipboard2 ;store selected reference from clipboard into variable
			AltClipboard2 := 
			} 
		Clipboard := ClipSaved ;restore previous clipboard

	;##########Step 2 - Evaluate reference
		; reset variables
		FoundPosNonDigit := "" 
		FoundPosNonDigit2 := ""
		FoundPosDigit := ""
		vEndNonDigit := ""
		vEndDigit := ""
		vEnd := ""
		vEndOriginal := ""

		; check for non-digit end in reference
		FoundPosNonDigit := RegExMatch(vReferenceOriginal, "([a-zA-Z]+)$", vEndNonDigit) ;does string end with non-digit-group ("abc")?
		FoundPosNonDigit2 := RegExMatch(vEndNonDigit, "(\D+)$", vEndNonDigit) ;does string end with non-digit-alphabetical-group? needed to account for "." or "-"
		
		; check for digit end in reference
		FoundPosDigit := RegExMatch(vReferenceOriginal, "(\d+)$", vEndDigit) ;does string end with digit-group ("123")?

	;##########Step 3 - Declare variables needed for navigation in next steps

		If FoundPosNonDigit2 != 0 ; did we identify a non-digit-alphabetical reference end?
			{
			vEnd = %vEndNonDigit% ; declare the reference end to work with
			vEndOriginal = %vEndNonDigit%
			StringTrimRight, vReferenceOriginalTrimmed, vReferenceOriginal, StrLen(vEndOriginal) ;trim original reference by length of identified reference end
			return "CaseNonDigit" ; declare that we deal with a non-digit-alphabetical reference end
			} else if FoundPosDigit != 0 ; did we identify a digit reference end?
				{
				vEnd = %vEndDigit% ; declare the reference end to work with
				vEndOriginal = %vEndDigit%
				StringTrimRight, vReferenceOriginalTrimmed, vReferenceOriginal, StrLen(vEndOriginal) ;trim original reference by length of identified reference end
				return "CaseDigit" ; declare that we deal with a digit reference end
				} else { ; else return an error message
					MsgBox, 16, , % "Invalid input. Please select/enter string ending with digit or non-digit character." 
					return "Error"
					}
	return "Error"
	}

	
CaseDigit_Function(){

	global
	
	;##########Step 3a - Work with digit reference
	; CaseDigit:
	
		; If vDebug = 1 
		; {
		; MsgBox, 64, Debug, 
		; ( LTrim
		; CaseDigit_Function started.
		; )
		; }
	
	If (vModus != 1) 
		{	
		vEndNew = %vEnd%
		vEndNew++
		} else if (vModus = 1) {
			 vEndNew = %vEnd%
			 vEndNew--
			 vEndNew--
			 }
	
	;check for leading Zero
		FoundLeadingZeros := ""
		LeadingZeros := ""
		StringTrimRight, LeadingZeroTest, vEndOriginal, StrLen(vEndNew)
		FoundLeadingZeros := RegExMatch(LeadingZeroTest, "^0+", LeadingZeros)
		If vDebug = 1 
		{
		MsgBox, 64, Debug,
			( LTrim
			CaseDigit_Function: Debug: Checking for leading zeros
			FoundLeadingZeros: %FoundLeadingZeros%
			LeadingZeros: %LeadingZeros%
			)
		}

		vReferenceNew := vReferenceOriginalTrimmed LeadingZeros vEndNew ;join shortened reference with new reference end
		vReferenceOld := vReferenceNew
			
		
		If vDebug = 1 
		{
		MsgBox, 64, Debug, 
		( LTrim
		CaseDigit_Function: Debug
		vModus: %vModus%
		vCase: %vCase%
		vEnd: %vEnd%
		vEndNew: %vEndNew%
		vReferenceNew: %vReferenceNew%
		)
		}
		
		If (vModus = 2) 
		{
		SendInput {Backspace}%vReferenceNew% ; e.g., Backspace and then "900-0003"
		vEnd := vEndNew
		vModus := ""
		} 
		return
		
		If (vModus = 1) 
		{
		vEndNewTemp := vEndNew
		vEndNewTemp++
		vReferenceNewTemp := vReferenceOriginalTrimmed vEndNewTemp ;join shortened reference with new reference end
		
		SplashTextOn ,200 ,75 , Info, Next reference will be:`n%vReferenceNewTemp% 
		Sleep 5000
		SplashTextOff
		
		vReferenceNewTemp := ""
		vEnd := vEndNew
		vModus := ""
		} 
	}

CaseNonDigit_Function(){

	global
	
		; If vDebug = 1 
		; {
		; MsgBox, 64, Debug, 
		; ( LTrim
		; CaseNonDigit_Function started.
		; )
		; }

	;##########Step 3b - Work with non-digit reference	
	;CaseNonDigit:
		
		If StrLen(vEnd) > 2 
			{
			MsgBox 16, , Invalid input. Please enter/select string ending with max 2 non-digit characters (i.e., a - z, aa - zz).
			} else if vEnd is upper 
				{
				ASCIIcount = 64
				} else if vEnd is lower 
					{
					ASCIIcount = 96
					} else {
						Stringlower, vEnd, vEnd
						ASCIIcount = 96
					}

		NonDigitArray := Array()
		vLetter :=
		vLetter2 :=
		
		Loop, 26 {
			vLetter := Chr(A_Index + ASCIIcount)
			 NonDigitArray.Push(vLetter)
			 }
			 
		Loop 26 {
			vLetter2:= Chr(A_Index + ASCIIcount)
			Loop, 26 {
				vLetter := vLetter2 Chr(A_Index + ASCIIcount)
				NonDigitArray.Push(vLetter)
				}
			}	
		
		vEndNonDigitCounter := HasVal(NonDigitArray, vEnd)

		CaseNonDigitRepeat_Function()
	return
	}
		
CaseNonDigitRepeat_Function(){

	global
	
		; If vDebug = 1 
		; {
		; MsgBox, 64, Debug, 
		; ( LTrim
		; CaseNonDigitRepeat_Function started.
		; )
		; }
		
		If HasVal(NonDigitArray, vEnd) = 0 {

			vEndNonDigitCounter := 1
			
				SplashTextOn ,200 ,75 , Info, Invalid input. No valid reference ending found (i.e., a - z, aa - zz). Next reference ending will be:`n NonDigitArray[1]
				Sleep 500
				SplashTextOff
			
		} 
		else if HasVal(NonDigitArray, vEnd) = 702 {
		
			vEndNonDigitCounter := 1
			
				SplashTextOn ,200 ,75 , Info, Maximum of non-digit counter reached ("zz"). Next reference ending will be:`n NonDigitArray[1]
				Sleep 500
				SplashTextOff
		}			
		else if (vModus != 1) {
			vEndNonDigitCounter++
		} else if (vModus = 1) {
			 vEndNonDigitCounter--
			 vEndNonDigitCounter--
			 }
			
		vEndNew := NonDigitArray[vEndNonDigitCounter]
		
		vReferenceNew := vReferenceOriginalTrimmed vEndNew ;join shortened reference with new reference end
		vReferenceOld := vReferenceNew
		
		If vDebug = 1 
		{
		MsgBox, 64, Debug, 
		( LTrim
		CaseNonDigitRepeat_Function: Debug
		vModus: %vModus%
		vCase: %vCase%
		vEnd: %vEnd%
		vEndNew: %vEndNew%
		vEndNonDigitCounter: %vEndNonDigitCounter%
		vReferenceNew: %vReferenceNew%
		)
		}

		If (vModus = 2) 
		{
		SendInput {Backspace}%vReferenceNew% ; e.g., Backspace and then "900-0003"
		vEnd := vEndNew
		vModus := ""
		} 
		
		If (vModus = 1) 
		{
		vEndNonDigitCounterTemp := vEndNonDigitCounter
		vEndNonDigitCounterTemp++
		vEndNewTemp := NonDigitArray[vEndNonDigitCounterTemp]
		vReferenceNewTemp := vReferenceOriginalTrimmed vEndNewTemp ;join shortened reference with new reference end
		
		SplashTextOn ,200 ,75 , Info, Next reference will be:`n%vReferenceNewTemp% 
		Sleep 5000
		SplashTextOff
		
		vReferenceNewTemp := ""
		vEnd := vEndNew
		vModus := ""
		} 
	
	return
	}

HasVal(haystack, needle){
	for index, value in haystack
		if (value = needle)
			return index
	if !(IsObject(haystack))
		throw Exception("Bad haystack!", -1, haystack)
	return 0
	}
	
CaseReset_Function(){

	global
	
	If vDebug = 1 
	{
	MsgBox, 64, Debug, 
	( LTrim
	CaseReset_Function started.
	)
	}

	If (vCase = "CaseDigit") 
	{
	;vEnd--
	CaseDigit_Function()
	} 
	Else if (vCase = "CaseNonDigit") 
	{
	
		If (vEndNonDigitCounter = 1) {
		vEndNonDigitCounter := NonDigitArray.Length()
		}
		
		;vEndNonDigitCounter--
	
		CaseNonDigitRepeat_Function()	
	}
	
	; SplashTextOn ,200 ,75 , Info, Next reference will be:`n%vReferenceNew%
	; Sleep 500
	; SplashTextOff
	
		; If (vCase = "CaseDigit") 
	; {
	; vEnd--
	; CaseDigit_Function()
	; } 
	; Else if (vCase = "CaseNonDigit") 
	; {
	
		; If (vEndNonDigitCounter = 1) {
		; vEndNonDigitCounter := NonDigitArray.Length()
		; }
		
		; vEndNonDigitCounter--
	
		; CaseNonDigitRepeat_Function()	
	; }
	
	vEnd := vEndNew
	vModus := "" 
	return		
}
	
MenuHandler:
Run https://github.com/flw-collab/referencehelper/
return