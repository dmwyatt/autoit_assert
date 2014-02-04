#include-once
#include <Const.au3>


Func _Assertion($condition, $msg, $title=$PROGRAM_TITLE, $terminate=True)
	Local $bCondition = Execute($condition)
	If Not $bCondition Then
		_ErrMsgBox($msg, $title, $terminate)
	EndIf
	Return SetError(1)
EndFunc

Func _AssertStringContains($string, $substring, $msg, $title=$PROGRAM_TITLE, $terminate=True)
	$assertion = StringFormat("StringInStr('%s', '%s')", $string, $substring)

	Return _Assertion($assertion, $msg, $title, $terminate)
EndFunc

Func _AssertArrayLength($array, $length, $msg, $title=$PROGRAM_TITLE, $terminate=True)
	If Not UBound($array) = $length Then
		_ErrMsgBox($msg, $title, $terminate)
		Return SetError(1)
	EndIf
EndFunc

Func _AssertValueIsIn($value, $possible_matches, $msg, $title=$PROGRAM_TITLE, $terminate=True)
	;possible_matches can be an array of strings/ints/floats or a string of comma-separated values with no spaces
	If IsString($possible_matches) Then
		$aPossibles = StringSplit($possible_matches, ",")
	Else
		$aPossibles = $possible_matches
	EndIf

	$assertion = String(_ValueIsIn($value, $aPossibles))
	Return _Assertion($assertion, $msg, $title, $terminate)
EndFunc

Func _AssertValueIsNot($value, $not, $msg, $title=$PROGRAM_TITLE, $terminate=True)
	If VarGetType($value) <> VarGetType($not) Then
		; Different types, not the same
		Return
	EndIf
	$assertion = StringFormat("%s <> %s", $value, $not)

	Return _Assertion($assertion, $msg, $title, $terminate)
EndFunc

Func _AssertIsInt($value, $msg, $title=$PROGRAM_TITLE, $terminate=True)
	If Not IsInt($value) Then
		_ErrMsgBox($msg, $title, $terminate)
	EndIf
	Return SetError(1)
EndFunc

Func _AssertFileExists($fn, $msg, $title=$PROGRAM_TITLE, $terminate=True)
	$iFileExists = FileExists($fn)
	If $iFileExists = 0 Then
		_ErrMsgBox($msg, $title, $terminate)
	EndIf
	Return SetError(1)
EndFunc

Func _ValueIsIn($value, $possible_matches)
	For $match In $possible_matches
		If $value = $match Then Return True
	Next
	Return False
EndFunc

Func _ErrMsgBox($msg, $title=$PROGRAM_TITLE, $terminate=True)
	MsgBox(16 + 262144, $title, $msg)
	If $terminate Then
		Exit
	EndIf
EndFunc