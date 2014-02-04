#include-once


; #FUNCTION# ====================================================================================================================
; Name ..........: _Assertion
; Description ...: Evalutes a condition and displays a MsgBox if the result is False-y
; Syntax ........: _Assertion($sCondition, $sMsg, $sTitle[, $fTerminate = True])
; Parameters ....: $sCondition          - A string value containing an evalutable AutoIt expression.
;                  $sMsg                - A string value containing a message to display if $sCondition is False.
;                  $sTitle              - A string value containing the title of the MsgBox to display.
;                  $fTerminate          - [optional] A boolean value indicating whether the script should
;										  terminate on failure of assertion.
; Return values .: Returns 1 if failure and sets @error to 1.
; Author ........: Dustin Wyatt
; Modified ......:
; Remarks .......: Mostly meant to be used by more specialized assertion functions.
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Assertion($sCondition, $sMsg, $sTitle, $fTerminate=True)
	Local $bCondition = Execute($sCondition)
	If Not $bCondition Then
		_ErrMsgBox($sMsg, $sTitle, $fTerminate)
	EndIf
	Return SetError(1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssertStringContains
; Description ...: Tests that a string contains a string.
; Syntax ........: _AssertStringContains($sString, $sSubString, $sMsg, $sTitle[, $fTerminate = True])
; Parameters ....: $sString             - A string value to search within.
;                  $sSubString          - A string value to search.
;                  $sMsg                - A string value.  See _Assertion().
;                  $sTitle              - A string value.  See _Assertion().
;                  $fTerminate          - [optional] A boolean value. Default is True.  See _Assertion().
; Return values .: See _Assertion().
; Author ........: Dustin Wyatt
; Modified ......:
; Remarks .......: Note that this is NOT case sensitive!
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AssertStringContains($sString, $sSubString, $sMsg, $sTitle, $fTerminate=True)
	$sAssertion = StringFormat("StringInStr('%s', '%s')", $sString, $sSubString)

	Return _Assertion($sAssertion, $sMsg, $sTitle, $fTerminate)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssertArrayLength
; Description ...: Tests that an Array is a length.
; Syntax ........: _AssertArrayLength($aArray, $iLength, $sMsg, $sTitle[, $fTerminate = True])
; Parameters ....: $aArray              - An array of unknowns.
;                  $iLength             - An integer value.
;                  $sMsg                - A string value.  See _Assertion().
;                  $sTitle              - A string value.  See _Assertion().
;                  $fTerminate          - [optional] A boolean value. Default is True.  See _Assertion().
; Return values .: Returns 1 if failure and sets @error to 1.
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AssertArrayLength($aArray, $iLength, $sMsg, $sTitle, $fTerminate=True)
	If Not UBound($aArray) = $iLength Then
		_ErrMsgBox($sMsg, $sTitle, $fTerminate)
		Return SetError(1)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssertValueIsIn
; Description ...: Tests that an list contains a value.
; Syntax ........: _AssertValueIsIn($vValue, $vPossibleMatches, $sMsg, $sTitle[, $fTerminate = True])
; Parameters ....: $vValue              - A variant value.  This is the value we want to be in $vPossibleMatches.
;                  $vPossibleMatches    - A variant value. This is either an array, or a string containing comma-separated values.
;                  $sMsg                - A string value.  See _Assertion().
;                  $sTitle              - A string value.  See _Assertion().
;                  $fTerminate          - [optional] A boolean value. Default is True.  See _Assertion().
; Return values .: See _Assertion().
; Author ........: Your Name
; Modified ......:
; Remarks .......: $vPossibleMatches is either an array or a string like: "these,are,some,values"
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AssertValueIsIn($vValue, $vPossibleMatches, $sMsg, $sTitle, $fTerminate=True)
	;possible_matches can be an array of strings/ints/floats or a string of comma-separated values with no spaces
	If IsString($vPossibleMatches) Then
		$aPossibles = StringSplit($vPossibleMatches, ",")
	Else
		$aPossibles = $vPossibleMatches
	EndIf

	$sAssertion = String(_ValueIsIn($vValue, $aPossibles))
	Return _Assertion($sAssertion, $sMsg, $sTitle, $fTerminate)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssertValueIsNot
; Description ...: Tests that a value is not the same as another value.
; Syntax ........: _AssertValueIsNot($vValue, $vNot, $sMsg, $sTitle[, $fTerminate = True])
; Parameters ....: $vValue              - A variant value.
;                  $vNot                - A variant value. The value to test against.
;                  $sMsg                - A string value.  See _Assertion().
;                  $sTitle              - A string value.  See _Assertion().
;                  $fTerminate          - [optional] A boolean value. Default is True.  See _Assertion().
; Return values .: See _Assertion().
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AssertValueIsNot($vValue, $vNot, $sMsg, $sTitle, $fTerminate=True)
	If VarGetType($vValue) <> VarGetType($vNot) Then
		; Different types, not the same
		Return
	EndIf
	$sAssertion = StringFormat("%s <> %s", $vValue, $vNot)

	Return _Assertion($sAssertion, $sMsg, $sTitle, $fTerminate)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssertValueIs
; Description ...: Tests that a value equals another value.
; Syntax ........: _AssertValueIs($vValue, $vIs, $sMsg, $sTitle[, $fTerminate = True])
; Parameters ....: $vValue              - A variant value.
;                  $vIs                 - A variant value.
;                  $sMsg                - A string value.  See _Assertion().
;                  $sTitle              - A string value.  See _Assertion().
;                  $fTerminate          - [optional] A boolean value. Default is True.  See _Assertion().
; Return values .: See _Assertion().
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AssertValueIs($vValue, $vIs, $sMsg, $sTitle, $fTerminate=True)
	If VarGetType($vValue) <> VarGetType($vIs) Then
		; Different types, not the same
		_ErrMsgBox($sMsg, $sTitle, $fTerminate)
		Return SetError(1)
	EndIf
	$sAssertion = StringFormat("%s = %s", $vValue, $vIs)

	Return _Assertion($sAssertion, $sMsg, $sTitle, $fTerminate)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssertIsInt
; Description ...: Tests that a value is an integer.
; Syntax ........: _AssertIsInt($iValue, $sMsg, $sTitle[, $fTerminate = True])
; Parameters ....: $iValue              - An integer (hopefully) value.
;                  $sMsg                - A string value.  See _Assertion().
;                  $sTitle              - A string value.  See _Assertion().
;                  $fTerminate          - [optional] A boolean value. Default is True.  See _Assertion().
; Return values .: See _Assertion().
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AssertIsInt($iValue, $sMsg, $sTitle, $fTerminate=True)
	If Not IsInt($iValue) Then
		_ErrMsgBox($sMsg, $sTitle, $fTerminate)
	EndIf
	Return SetError(1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssertFileExists
; Description ...: Tests that a file exists.
; Syntax ........: _AssertFileExists($sFn, $sMsg, $sTitle[, $fTerminate = True])
; Parameters ....: $sFn                 - A string value.
;                  $sMsg                - A string value.  See _Assertion().
;                  $sTitle              - A string value.  See _Assertion().
;                  $fTerminate          - [optional] A boolean value. Default is True.  See _Assertion().
; Return values .: See _Assertion().
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AssertFileExists($sFn, $sMsg, $sTitle, $fTerminate=True)
	$iFileExists = FileExists($sFn)
	If $iFileExists = 0 Then
		_ErrMsgBox($sMsg, $sTitle, $fTerminate)
	EndIf
	Return SetError(1)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssertWinWithTitleExists
; Description ...: Tests that a window exists
; Syntax ........: _AssertWinWithTitleExists($sWinTitle, $sMsg, $sTitle[, $fTerminate = True[, $iTitleMatchMode = 1]])
; Parameters ....: $sWinTitle           - A string value. The title of the window.  See Remarks.
;                  $sMsg                - A string value.  See _Assertion().
;                  $sTitle              - A string value.  See _Assertion().
;                  $fTerminate          - [optional] A boolean value. Default is True.  See _Assertion().
;                  $iTitleMatchMode     - [optional] An integer value. Default is 1.  See WinTitleMatchMode in AutoIt docs.
; Return values .: See _Assertion().
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _AssertWinWithTitleExists($sWinTitle, $sMsg, $sTitle, $fTerminate=True, $iTitleMatchMode=1)
	$iPreviousSettings = Opt("WinTitleMatchMode", $iTitleMatchMode)
	$iExists = WinExists($sWinTitle)
	Opt("WinTitleMatchMode", $iPreviousSettings)
	_AssertValueIs($iExists, 1, $sMsg, $sTitle, $fTerminate)
EndFunc

Func _ValueIsIn($value, $possible_matches)
	For $match In $possible_matches
		If $value = $match Then Return True
	Next
	Return False
EndFunc

Func _ErrMsgBox($sMsg, $sTitle, $fTerminate=True)
	MsgBox(16 + 262144, $sTitle, $sMsg)
	If $fTerminate Then
		Exit
	EndIf
EndFunc