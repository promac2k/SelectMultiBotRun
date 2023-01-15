#RequireAdmin
Global Const $STDERR_MERGED = 8
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $NUMBER_DOUBLE = 3
Global Const $WINDOWS_NOONTOP = 0
Global Const $WINDOWS_ONTOP = 1
Global Const $MB_OK = 0
Global Const $MB_YESNO = 4
Global Const $IDYES = 6
Global Const $STR_ENTIRESPLIT = 1
Global Const $STR_NOCOUNT = 2
Global Const $_ARRAYCONSTANT_SORTINFOSIZE = 11
Global $__g_aArrayDisplay_SortInfo[$_ARRAYCONSTANT_SORTINFOSIZE]
Global Const $_ARRAYCONSTANT_tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" & "int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
#Au3Stripper_Ignore_Funcs=__ArrayDisplay_SortCallBack
Func __ArrayDisplay_SortCallBack($nItem1, $nItem2, $hWnd)
If $__g_aArrayDisplay_SortInfo[3] = $__g_aArrayDisplay_SortInfo[4] Then
If Not $__g_aArrayDisplay_SortInfo[7] Then
$__g_aArrayDisplay_SortInfo[5] *= -1
$__g_aArrayDisplay_SortInfo[7] = 1
EndIf
Else
$__g_aArrayDisplay_SortInfo[7] = 1
EndIf
$__g_aArrayDisplay_SortInfo[6] = $__g_aArrayDisplay_SortInfo[3]
Local $sVal1 = __ArrayDisplay_GetItemText($hWnd, $nItem1, $__g_aArrayDisplay_SortInfo[3])
Local $sVal2 = __ArrayDisplay_GetItemText($hWnd, $nItem2, $__g_aArrayDisplay_SortInfo[3])
If $__g_aArrayDisplay_SortInfo[8] = 1 Then
If(StringIsFloat($sVal1) Or StringIsInt($sVal1)) Then $sVal1 = Number($sVal1)
If(StringIsFloat($sVal2) Or StringIsInt($sVal2)) Then $sVal2 = Number($sVal2)
EndIf
Local $nResult
If $__g_aArrayDisplay_SortInfo[8] < 2 Then
$nResult = 0
If $sVal1 < $sVal2 Then
$nResult = -1
ElseIf $sVal1 > $sVal2 Then
$nResult = 1
EndIf
Else
$nResult = DllCall('shlwapi.dll', 'int', 'StrCmpLogicalW', 'wstr', $sVal1, 'wstr', $sVal2)[0]
EndIf
$nResult = $nResult * $__g_aArrayDisplay_SortInfo[5]
Return $nResult
EndFunc
Func __ArrayDisplay_GetItemText($hWnd, $iIndex, $iSubItem = 0)
Local $tBuffer = DllStructCreate("wchar Text[4096]")
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagLVITEM)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "TextMax", 4096)
DllStructSetData($tItem, "Text", $pBuffer)
If IsHWnd($hWnd) Then
DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x1073, "wparam", $iIndex, "struct*", $tItem)
Else
Local $pItem = DllStructGetPtr($tItem)
GUICtrlSendMsg($hWnd, 0x1073, $iIndex, $pItem)
EndIf
Return DllStructGetData($tBuffer, "Text")
EndFunc
Global Enum $ARRAYFILL_FORCE_DEFAULT, $ARRAYFILL_FORCE_SINGLEITEM, $ARRAYFILL_FORCE_INT, $ARRAYFILL_FORCE_NUMBER, $ARRAYFILL_FORCE_PTR, $ARRAYFILL_FORCE_HWND, $ARRAYFILL_FORCE_STRING, $ARRAYFILL_FORCE_BOOLEAN
Global Enum $ARRAYUNIQUE_NOCOUNT, $ARRAYUNIQUE_COUNT
Global Enum $ARRAYUNIQUE_AUTO, $ARRAYUNIQUE_FORCE32, $ARRAYUNIQUE_FORCE64, $ARRAYUNIQUE_MATCH, $ARRAYUNIQUE_DISTINCT
Func _ArrayInsert(ByRef $aArray, $vRange, $vValue = "", $iStart = 0, $sDelim_Item = "|", $sDelim_Row = @CRLF, $iForce = $ARRAYFILL_FORCE_DEFAULT)
If $vValue = Default Then $vValue = ""
If $iStart = Default Then $iStart = 0
If $sDelim_Item = Default Then $sDelim_Item = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iForce = Default Then $iForce = $ARRAYFILL_FORCE_DEFAULT
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
Local $hDataType = 0
Switch $iForce
Case $ARRAYFILL_FORCE_INT
$hDataType = Int
Case $ARRAYFILL_FORCE_NUMBER
$hDataType = Number
Case $ARRAYFILL_FORCE_PTR
$hDataType = Ptr
Case $ARRAYFILL_FORCE_HWND
$hDataType = Hwnd
Case $ARRAYFILL_FORCE_STRING
$hDataType = String
EndSwitch
Local $aSplit_1, $aSplit_2
If IsArray($vRange) Then
If UBound($vRange, $UBOUND_DIMENSIONS) <> 1 Or UBound($vRange, $UBOUND_ROWS) < 2 Then Return SetError(4, 0, -1)
Else
Local $iNumber
$vRange = StringStripWS($vRange, 8)
$aSplit_1 = StringSplit($vRange, ";")
$vRange = ""
For $i = 1 To $aSplit_1[0]
If Not StringRegExp($aSplit_1[$i], "^\d+(-\d+)?$") Then Return SetError(3, 0, -1)
$aSplit_2 = StringSplit($aSplit_1[$i], "-")
Switch $aSplit_2[0]
Case 1
$vRange &= $aSplit_2[1] & ";"
Case 2
If Number($aSplit_2[2]) >= Number($aSplit_2[1]) Then
$iNumber = $aSplit_2[1] - 1
Do
$iNumber += 1
$vRange &= $iNumber & ";"
Until $iNumber = $aSplit_2[2]
EndIf
EndSwitch
Next
$vRange = StringSplit(StringTrimRight($vRange, 1), ";")
EndIf
If $vRange[1] < 0 Or $vRange[$vRange[0]] > $iDim_1 Then Return SetError(5, 0, -1)
For $i = 2 To $vRange[0]
If $vRange[$i] < $vRange[$i - 1] Then Return SetError(3, 0, -1)
Next
Local $iCopyTo_Index = $iDim_1 + $vRange[0]
Local $iInsertPoint_Index = $vRange[0]
Local $iInsert_Index = $vRange[$iInsertPoint_Index]
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iForce = $ARRAYFILL_FORCE_SINGLEITEM Then
ReDim $aArray[$iDim_1 + $vRange[0] + 1]
For $iReadFromIndex = $iDim_1 To 0 Step -1
$aArray[$iCopyTo_Index] = $aArray[$iReadFromIndex]
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
$aArray[$iCopyTo_Index] = $vValue
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index < 1 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Return $iDim_1 + $vRange[0] + 1
EndIf
ReDim $aArray[$iDim_1 + $vRange[0] + 1]
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(5, 0, -1)
$hDataType = 0
Else
Local $aTmp = StringSplit($vValue, $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
If UBound($aTmp, $UBOUND_ROWS) = 1 Then
$aTmp[0] = $vValue
$hDataType = 0
EndIf
$vValue = $aTmp
EndIf
For $iReadFromIndex = $iDim_1 To 0 Step -1
$aArray[$iCopyTo_Index] = $aArray[$iReadFromIndex]
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
If $iInsertPoint_Index <= UBound($vValue, $UBOUND_ROWS) Then
If IsFunc($hDataType) Then
$aArray[$iCopyTo_Index] = $hDataType($vValue[$iInsertPoint_Index - 1])
Else
$aArray[$iCopyTo_Index] = $vValue[$iInsertPoint_Index - 1]
EndIf
Else
$aArray[$iCopyTo_Index] = ""
EndIf
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index = 0 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
If $iStart < 0 Or $iStart > $iDim_2 - 1 Then Return SetError(6, 0, -1)
Local $iValDim_1, $iValDim_2
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 2 Then Return SetError(7, 0, -1)
$iValDim_1 = UBound($vValue, $UBOUND_ROWS)
$iValDim_2 = UBound($vValue, $UBOUND_COLUMNS)
$hDataType = 0
Else
$aSplit_1 = StringSplit($vValue, $sDelim_Row, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iValDim_1 = UBound($aSplit_1, $UBOUND_ROWS)
StringReplace($aSplit_1[0], $sDelim_Item, "")
$iValDim_2 = @extended + 1
Local $aTmp[$iValDim_1][$iValDim_2]
For $i = 0 To $iValDim_1 - 1
$aSplit_2 = StringSplit($aSplit_1[$i], $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
For $j = 0 To $iValDim_2 - 1
$aTmp[$i][$j] = $aSplit_2[$j]
Next
Next
$vValue = $aTmp
EndIf
If UBound($vValue, $UBOUND_COLUMNS) + $iStart > UBound($aArray, $UBOUND_COLUMNS) Then Return SetError(8, 0, -1)
ReDim $aArray[$iDim_1 + $vRange[0] + 1][$iDim_2]
For $iReadFromIndex = $iDim_1 To 0 Step -1
For $j = 0 To $iDim_2 - 1
$aArray[$iCopyTo_Index][$j] = $aArray[$iReadFromIndex][$j]
Next
$iCopyTo_Index -= 1
$iInsert_Index = $vRange[$iInsertPoint_Index]
While $iReadFromIndex = $iInsert_Index
For $j = 0 To $iDim_2 - 1
If $j < $iStart Then
$aArray[$iCopyTo_Index][$j] = ""
ElseIf $j - $iStart > $iValDim_2 - 1 Then
$aArray[$iCopyTo_Index][$j] = ""
Else
If $iInsertPoint_Index - 1 < $iValDim_1 Then
If IsFunc($hDataType) Then
$aArray[$iCopyTo_Index][$j] = $hDataType($vValue[$iInsertPoint_Index - 1][$j - $iStart])
Else
$aArray[$iCopyTo_Index][$j] = $vValue[$iInsertPoint_Index - 1][$j - $iStart]
EndIf
Else
$aArray[$iCopyTo_Index][$j] = ""
EndIf
EndIf
Next
$iCopyTo_Index -= 1
$iInsertPoint_Index -= 1
If $iInsertPoint_Index = 0 Then ExitLoop 2
$iInsert_Index = $vRange[$iInsertPoint_Index]
WEnd
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return UBound($aArray, $UBOUND_ROWS)
EndFunc
Func _ArraySearch(Const ByRef $aArray, $vValue, $iStart = 0, $iEnd = 0, $iCase = 0, $iCompare = 0, $iForward = 1, $iSubItem = -1, $bRow = False)
If $iStart = Default Then $iStart = 0
If $iEnd = Default Then $iEnd = 0
If $iCase = Default Then $iCase = 0
If $iCompare = Default Then $iCompare = 0
If $iForward = Default Then $iForward = 1
If $iSubItem = Default Then $iSubItem = -1
If $bRow = Default Then $bRow = False
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray) - 1
If $iDim_1 = -1 Then Return SetError(3, 0, -1)
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
Local $bCompType = False
If $iCompare = 2 Then
$iCompare = 0
$bCompType = True
EndIf
If $bRow Then
If UBound($aArray, $UBOUND_DIMENSIONS) = 1 Then Return SetError(5, 0, -1)
If $iEnd < 1 Or $iEnd > $iDim_2 Then $iEnd = $iDim_2
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(4, 0, -1)
Else
If $iEnd < 1 Or $iEnd > $iDim_1 Then $iEnd = $iDim_1
If $iStart < 0 Then $iStart = 0
If $iStart > $iEnd Then Return SetError(4, 0, -1)
EndIf
Local $iStep = 1
If Not $iForward Then
Local $iTmp = $iStart
$iStart = $iEnd
$iEnd = $iTmp
$iStep = -1
EndIf
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If Not $iCompare Then
If Not $iCase Then
For $i = $iStart To $iEnd Step $iStep
If $bCompType And VarGetType($aArray[$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i] = $vValue Then Return $i
Next
Else
For $i = $iStart To $iEnd Step $iStep
If $bCompType And VarGetType($aArray[$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i] == $vValue Then Return $i
Next
EndIf
Else
For $i = $iStart To $iEnd Step $iStep
If $iCompare = 3 Then
If StringRegExp($aArray[$i], $vValue) Then Return $i
Else
If StringInStr($aArray[$i], $vValue, $iCase) > 0 Then Return $i
EndIf
Next
EndIf
Case 2
Local $iDim_Sub
If $bRow Then
$iDim_Sub = $iDim_1
If $iSubItem > $iDim_Sub Then $iSubItem = $iDim_Sub
If $iSubItem < 0 Then
$iSubItem = 0
Else
$iDim_Sub = $iSubItem
EndIf
Else
$iDim_Sub = $iDim_2
If $iSubItem > $iDim_Sub Then $iSubItem = $iDim_Sub
If $iSubItem < 0 Then
$iSubItem = 0
Else
$iDim_Sub = $iSubItem
EndIf
EndIf
For $j = $iSubItem To $iDim_Sub
If Not $iCompare Then
If Not $iCase Then
For $i = $iStart To $iEnd Step $iStep
If $bRow Then
If $bCompType And VarGetType($aArray[$j][$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$j][$i] = $vValue Then Return $i
Else
If $bCompType And VarGetType($aArray[$i][$j]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i][$j] = $vValue Then Return $i
EndIf
Next
Else
For $i = $iStart To $iEnd Step $iStep
If $bRow Then
If $bCompType And VarGetType($aArray[$j][$i]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$j][$i] == $vValue Then Return $i
Else
If $bCompType And VarGetType($aArray[$i][$j]) <> VarGetType($vValue) Then ContinueLoop
If $aArray[$i][$j] == $vValue Then Return $i
EndIf
Next
EndIf
Else
For $i = $iStart To $iEnd Step $iStep
If $iCompare = 3 Then
If $bRow Then
If StringRegExp($aArray[$j][$i], $vValue) Then Return $i
Else
If StringRegExp($aArray[$i][$j], $vValue) Then Return $i
EndIf
Else
If $bRow Then
If StringInStr($aArray[$j][$i], $vValue, $iCase) > 0 Then Return $i
Else
If StringInStr($aArray[$i][$j], $vValue, $iCase) > 0 Then Return $i
EndIf
EndIf
Next
EndIf
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return SetError(6, 0, -1)
EndFunc
Func _ArrayToString(Const ByRef $aArray, $sDelim_Col = "|", $iStart_Row = -1, $iEnd_Row = -1, $sDelim_Row = @CRLF, $iStart_Col = -1, $iEnd_Col = -1)
If $sDelim_Col = Default Then $sDelim_Col = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iStart_Row = Default Then $iStart_Row = -1
If $iEnd_Row = Default Then $iEnd_Row = -1
If $iStart_Col = Default Then $iStart_Col = -1
If $iEnd_Col = Default Then $iEnd_Col = -1
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS) - 1
If $iStart_Row = -1 Then $iStart_Row = 0
If $iEnd_Row = -1 Then $iEnd_Row = $iDim_1
If $iStart_Row < -1 Or $iEnd_Row < -1 Then Return SetError(3, 0, -1)
If $iStart_Row > $iDim_1 Or $iEnd_Row > $iDim_1 Then Return SetError(3, 0, "")
If $iStart_Row > $iEnd_Row Then Return SetError(4, 0, -1)
Local $sRet = ""
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
For $i = $iStart_Row To $iEnd_Row
$sRet &= $aArray[$i] & $sDelim_Col
Next
Return StringTrimRight($sRet, StringLen($sDelim_Col))
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS) - 1
If $iStart_Col = -1 Then $iStart_Col = 0
If $iEnd_Col = -1 Then $iEnd_Col = $iDim_2
If $iStart_Col < -1 Or $iEnd_Col < -1 Then Return SetError(5, 0, -1)
If $iStart_Col > $iDim_2 Or $iEnd_Col > $iDim_2 Then Return SetError(5, 0, -1)
If $iStart_Col > $iEnd_Col Then Return SetError(6, 0, -1)
For $i = $iStart_Row To $iEnd_Row
For $j = $iStart_Col To $iEnd_Col
$sRet &= $aArray[$i][$j] & $sDelim_Col
Next
$sRet = StringTrimRight($sRet, StringLen($sDelim_Col)) & $sDelim_Row
Next
Return StringTrimRight($sRet, StringLen($sDelim_Row))
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return 1
EndFunc
Func _ArrayUnique(Const ByRef $aArray, $iColumn = 0, $iBase = 0, $iCase = 0, $iCount = $ARRAYUNIQUE_COUNT, $iIntType = $ARRAYUNIQUE_AUTO)
If $iColumn = Default Then $iColumn = 0
If $iBase = Default Then $iBase = 0
If $iCase = Default Then $iCase = 0
If $iCount = Default Then $iCount = $ARRAYUNIQUE_COUNT
If UBound($aArray, $UBOUND_ROWS) = 0 Then Return SetError(1, 0, 0)
Local $iDims = UBound($aArray, $UBOUND_DIMENSIONS), $iNumColumns = UBound($aArray, $UBOUND_COLUMNS)
If $iDims > 2 Then Return SetError(2, 0, 0)
If $iBase < 0 Or $iBase > 1 Or(Not IsInt($iBase)) Then Return SetError(3, 0, 0)
If $iCase < 0 Or $iCase > 1 Or(Not IsInt($iCase)) Then Return SetError(3, 0, 0)
If $iCount < 0 Or $iCount > 1 Or(Not IsInt($iCount)) Then Return SetError(4, 0, 0)
If $iIntType < 0 Or $iIntType > 4 Or(Not IsInt($iIntType)) Then Return SetError(5, 0, 0)
If $iColumn < 0 Or($iNumColumns = 0 And $iColumn > 0) Or($iNumColumns > 0 And $iColumn >= $iNumColumns) Then Return SetError(6, 0, 0)
If $iIntType = $ARRAYUNIQUE_AUTO Then
Local $bInt, $sVarType
If $iDims = 1 Then
$bInt = IsInt($aArray[$iBase])
$sVarType = VarGetType($aArray[$iBase])
Else
$bInt = IsInt($aArray[$iBase][$iColumn])
$sVarType = VarGetType($aArray[$iBase][$iColumn])
EndIf
If $bInt And $sVarType = "Int64" Then
$iIntType = $ARRAYUNIQUE_FORCE64
Else
$iIntType = $ARRAYUNIQUE_FORCE32
EndIf
EndIf
ObjEvent("AutoIt.Error", __ArrayUnique_AutoErrFunc)
Local $oDictionary = ObjCreate("Scripting.Dictionary")
$oDictionary.CompareMode = Number(Not $iCase)
Local $vElem, $sType, $vKey, $bCOMError = False
For $i = $iBase To UBound($aArray) - 1
If $iDims = 1 Then
$vElem = $aArray[$i]
Else
$vElem = $aArray[$i][$iColumn]
EndIf
Switch $iIntType
Case $ARRAYUNIQUE_FORCE32
$oDictionary.Item($vElem)
If @error Then
$bCOMError = True
ExitLoop
EndIf
Case $ARRAYUNIQUE_FORCE64
$sType = VarGetType($vElem)
If $sType = "Int32" Then
$bCOMError = True
ExitLoop
EndIf
$vKey = "#" & $sType & "#" & String($vElem)
If Not $oDictionary.Item($vKey) Then
$oDictionary($vKey) = $vElem
EndIf
Case $ARRAYUNIQUE_MATCH
$sType = VarGetType($vElem)
If StringLeft($sType, 3) = "Int" Then
$vKey = "#Int#" & String($vElem)
Else
$vKey = "#" & $sType & "#" & String($vElem)
EndIf
If Not $oDictionary.Item($vKey) Then
$oDictionary($vKey) = $vElem
EndIf
Case $ARRAYUNIQUE_DISTINCT
$vKey = "#" & VarGetType($vElem) & "#" & String($vElem)
If Not $oDictionary.Item($vKey) Then
$oDictionary($vKey) = $vElem
EndIf
EndSwitch
Next
Local $aValues, $j = 0
If $bCOMError Then
Return SetError(7, 0, 0)
ElseIf $iIntType <> $ARRAYUNIQUE_FORCE32 Then
Local $aValues[$oDictionary.Count]
For $vKey In $oDictionary.Keys()
$aValues[$j] = $oDictionary($vKey)
If StringLeft($vKey, 5) = "#Ptr#" Then
$aValues[$j] = Ptr($aValues[$j])
EndIf
$j += 1
Next
Else
$aValues = $oDictionary.Keys()
EndIf
If $iCount Then
_ArrayInsert($aValues, 0, $oDictionary.Count)
EndIf
Return $aValues
EndFunc
Func __ArrayUnique_AutoErrFunc()
EndFunc
Global Const $LBS_NOTIFY = 0x00000001
Global Const $LBS_SORT = 0x00000002
Global Const $LBS_NOSEL = 0x00004000
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
Global Const $GUI_CHECKED = 1
Global Const $GUI_SHOW = 16
Global Const $GUI_HIDE = 32
Global Const $GUI_ENABLE = 64
Global Const $GUI_DISABLE = 128
Func _SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
Local $aResult = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
If @error Then Return SetError(@error, @extended, "")
If $iReturn >= 0 And $iReturn <= 4 Then Return $aResult[$iReturn]
Return $aResult
EndFunc
Global Const $_UDF_GlobalIDs_OFFSET = 2
Global Const $_UDF_GlobalID_MAX_WIN = 16
Global Const $_UDF_STARTID = 10000
Global Const $_UDF_GlobalID_MAX_IDS = 55535
Global Const $__UDFGUICONSTANT_WS_VISIBLE = 0x10000000
Global Const $__UDFGUICONSTANT_WS_CHILD = 0x40000000
Global $__g_aUDF_GlobalIDs_Used[$_UDF_GlobalID_MAX_WIN][$_UDF_GlobalID_MAX_IDS + $_UDF_GlobalIDs_OFFSET + 1]
Func __UDF_GetNextGlobalID($hWnd)
Local $nCtrlID, $iUsedIndex = -1, $bAllUsed = True
If Not WinExists($hWnd) Then Return SetError(-1, -1, 0)
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] <> 0 Then
If Not WinExists($__g_aUDF_GlobalIDs_Used[$iIndex][0]) Then
For $x = 0 To UBound($__g_aUDF_GlobalIDs_Used, $UBOUND_COLUMNS) - 1
$__g_aUDF_GlobalIDs_Used[$iIndex][$x] = 0
Next
$__g_aUDF_GlobalIDs_Used[$iIndex][1] = $_UDF_STARTID
$bAllUsed = False
EndIf
EndIf
Next
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] = $hWnd Then
$iUsedIndex = $iIndex
ExitLoop
EndIf
Next
If $iUsedIndex = -1 Then
For $iIndex = 0 To $_UDF_GlobalID_MAX_WIN - 1
If $__g_aUDF_GlobalIDs_Used[$iIndex][0] = 0 Then
$__g_aUDF_GlobalIDs_Used[$iIndex][0] = $hWnd
$__g_aUDF_GlobalIDs_Used[$iIndex][1] = $_UDF_STARTID
$bAllUsed = False
$iUsedIndex = $iIndex
ExitLoop
EndIf
Next
EndIf
If $iUsedIndex = -1 And $bAllUsed Then Return SetError(16, 0, 0)
If $__g_aUDF_GlobalIDs_Used[$iUsedIndex][1] = $_UDF_STARTID + $_UDF_GlobalID_MAX_IDS Then
For $iIDIndex = $_UDF_GlobalIDs_OFFSET To UBound($__g_aUDF_GlobalIDs_Used, $UBOUND_COLUMNS) - 1
If $__g_aUDF_GlobalIDs_Used[$iUsedIndex][$iIDIndex] = 0 Then
$nCtrlID =($iIDIndex - $_UDF_GlobalIDs_OFFSET) + 10000
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][$iIDIndex] = $nCtrlID
Return $nCtrlID
EndIf
Next
Return SetError(-1, $_UDF_GlobalID_MAX_IDS, 0)
EndIf
$nCtrlID = $__g_aUDF_GlobalIDs_Used[$iUsedIndex][1]
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][1] += 1
$__g_aUDF_GlobalIDs_Used[$iUsedIndex][($nCtrlID - 10000) + $_UDF_GlobalIDs_OFFSET] = $nCtrlID
Return $nCtrlID
EndFunc
Global Const $tagPOINT = "struct;long X;long Y;endstruct"
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagNMHDR = "struct;hwnd hWndFrom;uint_ptr IDFrom;INT Code;endstruct"
Global Const $tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" & "int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
Global Const $tagNMITEMACTIVATE = $tagNMHDR & ";int Index;int SubItem;uint NewState;uint OldState;uint Changed;" & $tagPOINT & ";lparam lParam;uint KeyFlags"
Global Const $tagMENUINFO = "dword Size;INT Mask;dword Style;uint YMax;handle hBack;dword ContextHelpID;ulong_ptr MenuData"
Global Const $tagMENUITEMINFO = "uint Size;uint Mask;uint Type;uint State;uint ID;handle SubMenu;handle BmpChecked;handle BmpUnchecked;" & "ulong_ptr ItemData;ptr TypeData;uint CCH;handle BmpItem"
Global Const $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
Global Const $__WINVER = __WINVER()
Func _WinAPI_GetDlgCtrlID($hWnd)
Local $aResult = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetModuleHandle($sModuleName)
Local $sModuleNameType = "wstr"
If $sModuleName = "" Then
$sModuleName = 0
$sModuleNameType = "ptr"
EndIf
Local $aResult = DllCall("kernel32.dll", "handle", "GetModuleHandleW", $sModuleNameType, $sModuleName)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func __WINVER()
Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))
Local $aRet = DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)
If @error Or Not $aRet[0] Then Return SetError(@error, @extended, 0)
Return BitOR(BitShift(DllStructGetData($tOSVI, 2), -8), DllStructGetData($tOSVI, 3))
EndFunc
Func _WinAPI_ScreenToClient($hWnd, ByRef $tPoint)
Local $aResult = DllCall("user32.dll", "bool", "ScreenToClient", "hwnd", $hWnd, "struct*", $tPoint)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_CloseHandle($hObject)
Local $aResult = DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hObject)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_GetMousePos($bToClient = False, $hWnd = 0)
Local $iMode = Opt("MouseCoordMode", 1)
Local $aPos = MouseGetPos()
Opt("MouseCoordMode", $iMode)
Local $tPoint = DllStructCreate($tagPOINT)
DllStructSetData($tPoint, "X", $aPos[0])
DllStructSetData($tPoint, "Y", $aPos[1])
If $bToClient And Not _WinAPI_ScreenToClient($hWnd, $tPoint) Then Return SetError(@error + 20, @extended, 0)
Return $tPoint
EndFunc
Func _WinAPI_GetMousePosX($bToClient = False, $hWnd = 0)
Local $tPoint = _WinAPI_GetMousePos($bToClient, $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return DllStructGetData($tPoint, "X")
EndFunc
Func _WinAPI_GetMousePosY($bToClient = False, $hWnd = 0)
Local $tPoint = _WinAPI_GetMousePos($bToClient, $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return DllStructGetData($tPoint, "Y")
EndFunc
Global $__g_aInProcess_WinAPI[64][2] = [[0, 0]]
Func _WinAPI_CreateWindowEx($iExStyle, $sClass, $sName, $iStyle, $iX, $iY, $iWidth, $iHeight, $hParent, $hMenu = 0, $hInstance = 0, $pParam = 0)
If $hInstance = 0 Then $hInstance = _WinAPI_GetModuleHandle("")
Local $aResult = DllCall("user32.dll", "hwnd", "CreateWindowExW", "dword", $iExStyle, "wstr", $sClass, "wstr", $sName, "dword", $iStyle, "int", $iX, "int", $iY, "int", $iWidth, "int", $iHeight, "hwnd", $hParent, "handle", $hMenu, "handle", $hInstance, "struct*", $pParam)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetWindowHeight($hWnd)
Local $tRECT = _WinAPI_GetWindowRect($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return DllStructGetData($tRECT, "Bottom") - DllStructGetData($tRECT, "Top")
EndFunc
Func _WinAPI_GetWindowRect($hWnd)
Local $tRECT = DllStructCreate($tagRECT)
Local $aRet = DllCall("user32.dll", "bool", "GetWindowRect", "hwnd", $hWnd, "struct*", $tRECT)
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, 0)
Return $tRECT
EndFunc
Func _WinAPI_GetWindowThreadProcessId($hWnd, ByRef $iPID)
Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error, @extended, 0)
$iPID = $aResult[2]
Return $aResult[0]
EndFunc
Func _WinAPI_GetWindowWidth($hWnd)
Local $tRECT = _WinAPI_GetWindowRect($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return DllStructGetData($tRECT, "Right") - DllStructGetData($tRECT, "Left")
EndFunc
Func _WinAPI_InProcess($hWnd, ByRef $hLastWnd)
If $hWnd = $hLastWnd Then Return True
For $iI = $__g_aInProcess_WinAPI[0][0] To 1 Step -1
If $hWnd = $__g_aInProcess_WinAPI[$iI][0] Then
If $__g_aInProcess_WinAPI[$iI][1] Then
$hLastWnd = $hWnd
Return True
Else
Return False
EndIf
EndIf
Next
Local $iPID
_WinAPI_GetWindowThreadProcessId($hWnd, $iPID)
Local $iCount = $__g_aInProcess_WinAPI[0][0] + 1
If $iCount >= 64 Then $iCount = 1
$__g_aInProcess_WinAPI[0][0] = $iCount
$__g_aInProcess_WinAPI[$iCount][0] = $hWnd
$__g_aInProcess_WinAPI[$iCount][1] =($iPID = @AutoItPID)
Return $__g_aInProcess_WinAPI[$iCount][1]
EndFunc
Func _WinAPI_MoveWindow($hWnd, $iX, $iY, $iWidth, $iHeight, $bRepaint = True)
Local $aResult = DllCall("user32.dll", "bool", "MoveWindow", "hwnd", $hWnd, "int", $iX, "int", $iY, "int", $iWidth, "int", $iHeight, "bool", $bRepaint)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _WinAPI_SetParent($hWndChild, $hWndParent)
Local $aResult = DllCall("user32.dll", "hwnd", "SetParent", "hwnd", $hWndChild, "hwnd", $hWndParent)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _WinAPI_GetLastError(Const $_iCurrentError = @error, Const $_iCurrentExtended = @extended)
Local $aResult = DllCall("kernel32.dll", "dword", "GetLastError")
Return SetError($_iCurrentError, $_iCurrentExtended, $aResult[0])
EndFunc
Global Const $WS_GROUP = 0x00020000
Global Const $WM_NOTIFY = 0x004E
Global Const $WM_CONTEXTMENU = 0x007B
Global Const $NM_FIRST = 0
Global Const $NM_DBLCLK = $NM_FIRST - 3
Global Const $CBS_AUTOHSCROLL = 0x40
Global Const $CBS_DROPDOWNLIST = 0x3
Func _WinAPI_GetTempFileName($sFilePath, $sPrefix = '')
Local $aRet = DllCall('kernel32.dll', 'uint', 'GetTempFileNameW', 'wstr', $sFilePath, 'wstr', $sPrefix, 'uint', 0, 'wstr', '')
If @error Or Not $aRet[0] Then Return SetError(@error + 10, @extended, '')
Return $aRet[4]
EndFunc
Global Const $PROCESS_VM_OPERATION = 0x00000008
Global Const $PROCESS_VM_READ = 0x00000010
Global Const $PROCESS_VM_WRITE = 0x00000020
Global Const $PROCESS_QUERY_INFORMATION = 0x00000400
Global Const $PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
Global Const $SE_PRIVILEGE_ENABLED = 0x00000002
Global Enum $SECURITYANONYMOUS = 0, $SECURITYIDENTIFICATION, $SECURITYIMPERSONATION, $SECURITYDELEGATION
Global Const $TOKEN_QUERY = 0x00000008
Global Const $TOKEN_ADJUST_PRIVILEGES = 0x00000020
Func _Security__AdjustTokenPrivileges($hToken, $bDisableAll, $tNewState, $iBufferLen, $tPrevState = 0, $pRequired = 0)
Local $aCall = DllCall("advapi32.dll", "bool", "AdjustTokenPrivileges", "handle", $hToken, "bool", $bDisableAll, "struct*", $tNewState, "dword", $iBufferLen, "struct*", $tPrevState, "struct*", $pRequired)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__ImpersonateSelf($iLevel = $SECURITYIMPERSONATION)
Local $aCall = DllCall("advapi32.dll", "bool", "ImpersonateSelf", "int", $iLevel)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__LookupPrivilegeValue($sSystem, $sName)
Local $aCall = DllCall("advapi32.dll", "bool", "LookupPrivilegeValueW", "wstr", $sSystem, "wstr", $sName, "int64*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[3]
EndFunc
Func _Security__OpenThreadToken($iAccess, $hThread = 0, $bOpenAsSelf = False)
If $hThread = 0 Then
Local $aResult = DllCall("kernel32.dll", "handle", "GetCurrentThread")
If @error Then Return SetError(@error + 10, @extended, 0)
$hThread = $aResult[0]
EndIf
Local $aCall = DllCall("advapi32.dll", "bool", "OpenThreadToken", "handle", $hThread, "dword", $iAccess, "bool", $bOpenAsSelf, "handle*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error, @extended, 0)
Return $aCall[4]
EndFunc
Func _Security__OpenThreadTokenEx($iAccess, $hThread = 0, $bOpenAsSelf = False)
Local $hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then
Local Const $ERROR_NO_TOKEN = 1008
If _WinAPI_GetLastError() <> $ERROR_NO_TOKEN Then Return SetError(20, _WinAPI_GetLastError(), 0)
If Not _Security__ImpersonateSelf() Then Return SetError(@error + 10, _WinAPI_GetLastError(), 0)
$hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then Return SetError(@error, _WinAPI_GetLastError(), 0)
EndIf
Return $hToken
EndFunc
Func _Security__SetPrivilege($hToken, $sPrivilege, $bEnable)
Local $iLUID = _Security__LookupPrivilegeValue("", $sPrivilege)
If $iLUID = 0 Then Return SetError(@error + 10, @extended, False)
Local Const $tagTOKEN_PRIVILEGES = "dword Count;align 4;int64 LUID;dword Attributes"
Local $tCurrState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iCurrState = DllStructGetSize($tCurrState)
Local $tPrevState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iPrevState = DllStructGetSize($tPrevState)
Local $tRequired = DllStructCreate("int Data")
DllStructSetData($tCurrState, "Count", 1)
DllStructSetData($tCurrState, "LUID", $iLUID)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tCurrState, $iCurrState, $tPrevState, $tRequired) Then Return SetError(2, @error, False)
DllStructSetData($tPrevState, "Count", 1)
DllStructSetData($tPrevState, "LUID", $iLUID)
Local $iAttributes = DllStructGetData($tPrevState, "Attributes")
If $bEnable Then
$iAttributes = BitOR($iAttributes, $SE_PRIVILEGE_ENABLED)
Else
$iAttributes = BitAND($iAttributes, BitNOT($SE_PRIVILEGE_ENABLED))
EndIf
DllStructSetData($tPrevState, "Attributes", $iAttributes)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tPrevState, $iPrevState, $tCurrState, $tRequired) Then Return SetError(3, @error, False)
Return True
EndFunc
Func _WinAPI_OpenProcess($iAccess, $bInherit, $iPID, $bDebugPriv = False)
Local $aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return $aResult[0]
If Not $bDebugPriv Then Return SetError(100, 0, 0)
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 10, @extended, 0)
_Security__SetPrivilege($hToken, "SeDebugPrivilege", True)
Local $iError = @error
Local $iExtended = @extended
Local $iRet = 0
If Not @error Then
$aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
$iError = @error
$iExtended = @extended
If $aResult[0] Then $iRet = $aResult[0]
_Security__SetPrivilege($hToken, "SeDebugPrivilege", False)
If @error Then
$iError = @error + 20
$iExtended = @extended
EndIf
Else
$iError = @error + 30
EndIf
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hToken)
Return SetError($iError, $iExtended, $iRet)
EndFunc
Func _WinAPI_WaitForSingleObject($hHandle, $iTimeout = -1)
Local $aResult = DllCall("kernel32.dll", "INT", "WaitForSingleObject", "handle", $hHandle, "dword", $iTimeout)
If @error Then Return SetError(@error, @extended, -1)
Return $aResult[0]
EndFunc
Func _WinAPI_GetVersion()
Return Number(BitAND(BitShift($__WINVER, 8), 0xFF) & '.' & BitAND($__WINVER, 0xFF), $NUMBER_DOUBLE)
EndFunc
Global Const $MF_STRING = 0x0
Global Const $MF_SEPARATOR = 0x00000800
Global Const $MFS_GRAYED = 0x00000003
Global Const $MFS_DISABLED = $MFS_GRAYED
Global Const $MFT_STRING = $MF_STRING
Global Const $MFT_SEPARATOR = $MF_SEPARATOR
Global Const $MIIM_STATE = 0x00000001
Global Const $MIIM_ID = 0x00000002
Global Const $MIIM_SUBMENU = 0x00000004
Global Const $MIIM_DATAMASK = 0x0000003F
Global Const $MIIM_STRING = 0x00000040
Global Const $MIIM_FTYPE = 0x00000100
Global Const $MIM_STYLE = 0x00000010
Global Const $MNS_CHECKORBMP = 0x04000000
Global Const $TPM_LEFTBUTTON = 0x0
Global Const $TPM_LEFTALIGN = 0x0
Global Const $TPM_TOPALIGN = 0x0
Global Const $TPM_RIGHTBUTTON = 0x00000002
Global Const $TPM_CENTERALIGN = 0x00000004
Global Const $TPM_RIGHTALIGN = 0x00000008
Global Const $TPM_VCENTERALIGN = 0x00000010
Global Const $TPM_BOTTOMALIGN = 0x00000020
Global Const $TPM_NONOTIFY = 0x00000080
Global Const $TPM_RETURNCMD = 0x00000100
Func _GUICtrlMenu_CreatePopup($iStyle = $MNS_CHECKORBMP)
Local $aResult = DllCall("user32.dll", "handle", "CreatePopupMenu")
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] = 0 Then Return SetError(10, 0, 0)
_GUICtrlMenu_SetMenuStyle($aResult[0], $iStyle)
Return $aResult[0]
EndFunc
Func _GUICtrlMenu_GetItemInfo($hMenu, $iItem, $bByPos = True)
Local $tInfo = DllStructCreate($tagMENUITEMINFO)
DllStructSetData($tInfo, "Size", DllStructGetSize($tInfo))
DllStructSetData($tInfo, "Mask", $MIIM_DATAMASK)
Local $aResult = DllCall("user32.dll", "bool", "GetMenuItemInfo", "handle", $hMenu, "uint", $iItem, "bool", $bByPos, "struct*", $tInfo)
If @error Then Return SetError(@error, @extended, 0)
Return SetExtended($aResult[0], $tInfo)
EndFunc
Func _GUICtrlMenu_GetItemStateEx($hMenu, $iItem, $bByPos = True)
Local $tInfo = _GUICtrlMenu_GetItemInfo($hMenu, $iItem, $bByPos)
Return DllStructGetData($tInfo, "State")
EndFunc
Func _GUICtrlMenu_InsertMenuItem($hMenu, $iIndex, $sText, $iCmdID = 0, $hSubMenu = 0)
Local $tMenu = DllStructCreate($tagMENUITEMINFO)
DllStructSetData($tMenu, "Size", DllStructGetSize($tMenu))
DllStructSetData($tMenu, "ID", $iCmdID)
DllStructSetData($tMenu, "SubMenu", $hSubMenu)
If $sText = "" Then
DllStructSetData($tMenu, "Mask", $MIIM_FTYPE)
DllStructSetData($tMenu, "Type", $MFT_SEPARATOR)
Else
DllStructSetData($tMenu, "Mask", BitOR($MIIM_ID, $MIIM_STRING, $MIIM_SUBMENU))
DllStructSetData($tMenu, "Type", $MFT_STRING)
Local $tText = DllStructCreate("wchar Text[" & StringLen($sText) + 1 & "]")
DllStructSetData($tText, "Text", $sText)
DllStructSetData($tMenu, "TypeData", DllStructGetPtr($tText))
EndIf
Local $aResult = DllCall("user32.dll", "bool", "InsertMenuItemW", "handle", $hMenu, "uint", $iIndex, "bool", True, "struct*", $tMenu)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _GUICtrlMenu_SetItemDisabled($hMenu, $iItem, $bState = True, $bByPos = True)
Return _GUICtrlMenu_SetItemState($hMenu, $iItem, BitOR($MFS_DISABLED, $MFS_GRAYED), $bState, $bByPos)
EndFunc
Func _GUICtrlMenu_SetItemEnabled($hMenu, $iItem, $bState = True, $bByPos = True)
Return _GUICtrlMenu_SetItemState($hMenu, $iItem, BitOR($MFS_DISABLED, $MFS_GRAYED), Not $bState, $bByPos)
EndFunc
Func _GUICtrlMenu_SetItemInfo($hMenu, $iItem, ByRef $tInfo, $bByPos = True)
DllStructSetData($tInfo, "Size", DllStructGetSize($tInfo))
Local $aResult = DllCall("user32.dll", "bool", "SetMenuItemInfoW", "handle", $hMenu, "uint", $iItem, "bool", $bByPos, "struct*", $tInfo)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _GUICtrlMenu_SetItemState($hMenu, $iItem, $iState, $bState = True, $bByPos = True)
Local $iFlag = _GUICtrlMenu_GetItemStateEx($hMenu, $iItem, $bByPos)
If $bState Then
$iState = BitOR($iFlag, $iState)
Else
$iState = BitAND($iFlag, BitNOT($iState))
EndIf
Local $tInfo = DllStructCreate($tagMENUITEMINFO)
DllStructSetData($tInfo, "Size", DllStructGetSize($tInfo))
DllStructSetData($tInfo, "Mask", $MIIM_STATE)
DllStructSetData($tInfo, "State", $iState)
Return _GUICtrlMenu_SetItemInfo($hMenu, $iItem, $tInfo, $bByPos)
EndFunc
Func _GUICtrlMenu_SetMenuInfo($hMenu, ByRef $tInfo)
DllStructSetData($tInfo, "Size", DllStructGetSize($tInfo))
Local $aResult = DllCall("user32.dll", "bool", "SetMenuInfo", "handle", $hMenu, "struct*", $tInfo)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _GUICtrlMenu_SetMenuStyle($hMenu, $iStyle)
Local $tInfo = DllStructCreate($tagMENUINFO)
DllStructSetData($tInfo, "Mask", $MIM_STYLE)
DllStructSetData($tInfo, "Style", $iStyle)
Return _GUICtrlMenu_SetMenuInfo($hMenu, $tInfo)
EndFunc
Func _GUICtrlMenu_TrackPopupMenu($hMenu, $hWnd, $iX = -1, $iY = -1, $iAlignX = 1, $iAlignY = 1, $iNotify = 0, $iButtons = 0)
If $iX = -1 Then $iX = _WinAPI_GetMousePosX()
If $iY = -1 Then $iY = _WinAPI_GetMousePosY()
Local $iFlags = 0
Switch $iAlignX
Case 1
$iFlags = BitOR($iFlags, $TPM_LEFTALIGN)
Case 2
$iFlags = BitOR($iFlags, $TPM_RIGHTALIGN)
Case Else
$iFlags = BitOR($iFlags, $TPM_CENTERALIGN)
EndSwitch
Switch $iAlignY
Case 1
$iFlags = BitOR($iFlags, $TPM_TOPALIGN)
Case 2
$iFlags = BitOR($iFlags, $TPM_VCENTERALIGN)
Case Else
$iFlags = BitOR($iFlags, $TPM_BOTTOMALIGN)
EndSwitch
If BitAND($iNotify, 1) <> 0 Then $iFlags = BitOR($iFlags, $TPM_NONOTIFY)
If BitAND($iNotify, 2) <> 0 Then $iFlags = BitOR($iFlags, $TPM_RETURNCMD)
Switch $iButtons
Case 1
$iFlags = BitOR($iFlags, $TPM_RIGHTBUTTON)
Case Else
$iFlags = BitOR($iFlags, $TPM_LEFTBUTTON)
EndSwitch
Local $aResult = DllCall("user32.dll", "bool", "TrackPopupMenu", "handle", $hMenu, "uint", $iFlags, "int", $iX, "int", $iY, "int", 0, "hwnd", $hWnd, "ptr", 0)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Global Const $INET_FORCERELOAD = 1
Global Const $INET_DOWNLOADBACKGROUND = 1
Global Const $INET_DOWNLOADCOMPLETE = 2
Func _IsPressed($sHexKey, $vDLL = "user32.dll")
Local $aReturn = DllCall($vDLL, "short", "GetAsyncKeyState", "int", "0x" & $sHexKey)
If @error Then Return SetError(@error, @extended, False)
Return BitAND($aReturn[0], 0x8000) <> 0
EndFunc
Func _VersionCompare($sVersion1, $sVersion2)
If $sVersion1 = $sVersion2 Then Return 0
Local $sSubVersion1 = "", $sSubVersion2 = ""
If StringIsAlpha(StringRight($sVersion1, 1)) Then
$sSubVersion1 = StringRight($sVersion1, 1)
$sVersion1 = StringTrimRight($sVersion1, 1)
EndIf
If StringIsAlpha(StringRight($sVersion2, 1)) Then
$sSubVersion2 = StringRight($sVersion2, 1)
$sVersion2 = StringTrimRight($sVersion2, 1)
EndIf
Local $aVersion1 = StringSplit($sVersion1, ".,"), $aVersion2 = StringSplit($sVersion2, ".,")
Local $iPartDifference =($aVersion1[0] - $aVersion2[0])
If $iPartDifference < 0 Then
ReDim $aVersion1[UBound($aVersion2)]
$aVersion1[0] = UBound($aVersion1) - 1
For $i =(UBound($aVersion1) - Abs($iPartDifference)) To $aVersion1[0]
$aVersion1[$i] = "0"
Next
ElseIf $iPartDifference > 0 Then
ReDim $aVersion2[UBound($aVersion1)]
$aVersion2[0] = UBound($aVersion2) - 1
For $i =(UBound($aVersion2) - Abs($iPartDifference)) To $aVersion2[0]
$aVersion2[$i] = "0"
Next
EndIf
For $i = 1 To $aVersion1[0]
If StringIsDigit($aVersion1[$i]) And StringIsDigit($aVersion2[$i]) Then
If Number($aVersion1[$i]) > Number($aVersion2[$i]) Then
Return SetExtended(2, 1)
ElseIf Number($aVersion1[$i]) < Number($aVersion2[$i]) Then
Return SetExtended(2, -1)
ElseIf $i = $aVersion1[0] Then
If $sSubVersion1 > $sSubVersion2 Then
Return SetExtended(3, 1)
ElseIf $sSubVersion1 < $sSubVersion2 Then
Return SetExtended(3, -1)
EndIf
EndIf
Else
If $aVersion1[$i] > $aVersion2[$i] Then
Return SetExtended(1, 1)
ElseIf $aVersion1[$i] < $aVersion2[$i] Then
Return SetExtended(1, -1)
EndIf
EndIf
Next
Return SetExtended(Abs($iPartDifference), 0)
EndFunc
Global Const $HGDI_ERROR = Ptr(-1)
Global Const $INVALID_HANDLE_VALUE = Ptr(-1)
Global Const $KF_EXTENDED = 0x0100
Global Const $KF_ALTDOWN = 0x2000
Global Const $KF_UP = 0x8000
Global Const $LLKHF_EXTENDED = BitShift($KF_EXTENDED, 8)
Global Const $LLKHF_ALTDOWN = BitShift($KF_ALTDOWN, 8)
Global Const $LLKHF_UP = BitShift($KF_UP, 8)
Global Const $ES_READONLY = 2048
Global Const $COLOR_RED = 0xFF0000
Global Const $COLOR_WHITE = 0xFFFFFF
Global Const $LV_ERR = -1
Global Const $LVCF_FMT = 0x0001
Global Const $LVCF_IMAGE = 0x0010
Global Const $LVCF_TEXT = 0x0004
Global Const $LVCF_WIDTH = 0x0002
Global Const $LVCFMT_BITMAP_ON_RIGHT = 0x1000
Global Const $LVCFMT_CENTER = 0x0002
Global Const $LVCFMT_COL_HAS_IMAGES = 0x8000
Global Const $LVCFMT_IMAGE = 0x0800
Global Const $LVCFMT_LEFT = 0x0000
Global Const $LVCFMT_RIGHT = 0x0001
Global Const $LVIF_IMAGE = 0x00000002
Global Const $LVIF_PARAM = 0x00000004
Global Const $LVIF_TEXT = 0x00000001
Global Const $LVIS_SELECTED = 0x0002
Global Const $LVS_REPORT = 0x0001
Global Const $LVS_SHOWSELALWAYS = 0x0008
Global Const $LVM_FIRST = 0x1000
Global Const $LVM_DELETEALLITEMS =($LVM_FIRST + 9)
Global Const $LVM_GETHEADER =($LVM_FIRST + 31)
Global Const $LVM_GETITEMA =($LVM_FIRST + 5)
Global Const $LVM_GETITEMW =($LVM_FIRST + 75)
Global Const $LVM_GETITEMCOUNT =($LVM_FIRST + 4)
Global Const $LVM_GETITEMRECT =($LVM_FIRST + 14)
Global Const $LVM_GETITEMSTATE =($LVM_FIRST + 44)
Global Const $LVM_GETITEMTEXTA =($LVM_FIRST + 45)
Global Const $LVM_GETITEMTEXTW =($LVM_FIRST + 115)
Global Const $LVM_GETNEXTITEM =($LVM_FIRST + 12)
Global Const $LVM_GETSELECTEDCOUNT =($LVM_FIRST + 50)
Global Const $LVM_GETUNICODEFORMAT = 0x2000 + 6
Global Const $LVM_INSERTCOLUMNA =($LVM_FIRST + 27)
Global Const $LVM_INSERTCOLUMNW =($LVM_FIRST + 97)
Global Const $LVM_INSERTITEMA =($LVM_FIRST + 7)
Global Const $LVM_INSERTITEMW =($LVM_FIRST + 77)
Global Const $LVM_SETCOLUMNA =($LVM_FIRST + 26)
Global Const $LVM_SETCOLUMNW =($LVM_FIRST + 96)
Global Const $LVM_SETITEMA =($LVM_FIRST + 6)
Global Const $LVM_SETITEMW =($LVM_FIRST + 76)
Global Const $LVNI_ABOVE = 0x0100
Global Const $LVNI_BELOW = 0x0200
Global Const $LVNI_TOLEFT = 0x0400
Global Const $LVNI_TORIGHT = 0x0800
Global Const $LVNI_ALL = 0x0000
Global Const $LVNI_CUT = 0x0004
Global Const $LVNI_DROPHILITED = 0x0008
Global Const $LVNI_FOCUSED = 0x0001
Global Const $LVNI_SELECTED = 0x0002
Global Const $MEM_COMMIT = 0x00001000
Global Const $MEM_RESERVE = 0x00002000
Global Const $PAGE_READWRITE = 0x00000004
Global Const $MEM_RELEASE = 0x00008000
Global Const $tagMEMMAP = "handle hProc;ulong_ptr Size;ptr Mem"
Func _MemFree(ByRef $tMemMap)
Local $pMemory = DllStructGetData($tMemMap, "Mem")
Local $hProcess = DllStructGetData($tMemMap, "hProc")
Local $bResult = _MemVirtualFreeEx($hProcess, $pMemory, 0, $MEM_RELEASE)
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
If @error Then Return SetError(@error, @extended, False)
Return $bResult
EndFunc
Func _MemInit($hWnd, $iSize, ByRef $tMemMap)
Local $aResult = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error + 10, @extended, 0)
Local $iProcessID = $aResult[2]
If $iProcessID = 0 Then Return SetError(1, 0, 0)
Local $iAccess = BitOR($PROCESS_VM_OPERATION, $PROCESS_VM_READ, $PROCESS_VM_WRITE)
Local $hProcess = __Mem_OpenProcess($iAccess, False, $iProcessID, True)
Local $iAlloc = BitOR($MEM_RESERVE, $MEM_COMMIT)
Local $pMemory = _MemVirtualAllocEx($hProcess, 0, $iSize, $iAlloc, $PAGE_READWRITE)
If $pMemory = 0 Then Return SetError(2, 0, 0)
$tMemMap = DllStructCreate($tagMEMMAP)
DllStructSetData($tMemMap, "hProc", $hProcess)
DllStructSetData($tMemMap, "Size", $iSize)
DllStructSetData($tMemMap, "Mem", $pMemory)
Return $pMemory
EndFunc
Func _MemRead(ByRef $tMemMap, $pSrce, $pDest, $iSize)
Local $aResult = DllCall("kernel32.dll", "bool", "ReadProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pSrce, "struct*", $pDest, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemWrite(ByRef $tMemMap, $pSrce, $pDest = 0, $iSize = 0, $sSrce = "struct*")
If $pDest = 0 Then $pDest = DllStructGetData($tMemMap, "Mem")
If $iSize = 0 Then $iSize = DllStructGetData($tMemMap, "Size")
Local $aResult = DllCall("kernel32.dll", "bool", "WriteProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pDest, $sSrce, $pSrce, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func _MemVirtualAllocEx($hProcess, $pAddress, $iSize, $iAllocation, $iProtect)
Local $aResult = DllCall("kernel32.dll", "ptr", "VirtualAllocEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iAllocation, "dword", $iProtect)
If @error Then Return SetError(@error, @extended, 0)
Return $aResult[0]
EndFunc
Func _MemVirtualFreeEx($hProcess, $pAddress, $iSize, $iFreeType)
Local $aResult = DllCall("kernel32.dll", "bool", "VirtualFreeEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iFreeType)
If @error Then Return SetError(@error, @extended, False)
Return $aResult[0]
EndFunc
Func __Mem_OpenProcess($iAccess, $bInherit, $iPID, $bDebugPriv = False)
Local $aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
If @error Then Return SetError(@error, @extended, 0)
If $aResult[0] Then Return $aResult[0]
If Not $bDebugPriv Then Return SetError(100, 0, 0)
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 10, @extended, 0)
_Security__SetPrivilege($hToken, "SeDebugPrivilege", True)
Local $iError = @error
Local $iExtended = @extended
Local $iRet = 0
If Not @error Then
$aResult = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
$iError = @error
$iExtended = @extended
If $aResult[0] Then $iRet = $aResult[0]
_Security__SetPrivilege($hToken, "SeDebugPrivilege", False)
If @error Then
$iError = @error + 20
$iExtended = @extended
EndIf
Else
$iError = @error + 30
EndIf
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hToken)
Return SetError($iError, $iExtended, $iRet)
EndFunc
Global $__g_hLVLastWnd
Global Const $__LISTVIEWCONSTANT_SORTINFOSIZE = 11
Global $__g_aListViewSortInfo[1][$__LISTVIEWCONSTANT_SORTINFOSIZE]
Global Const $__LISTVIEWCONSTANT_WM_SETREDRAW = 0x000B
Global Const $tagLVCOLUMN = "uint Mask;int Fmt;int CX;ptr Text;int TextMax;int SubItem;int Image;int Order;int cxMin;int cxDefault;int cxIdeal"
Func _GUICtrlListView_AddItem($hWnd, $sText, $iImage = -1, $iParam = 0)
Return _GUICtrlListView_InsertItem($hWnd, $sText, -1, $iImage, $iParam)
EndFunc
Func _GUICtrlListView_AddSubItem($hWnd, $iIndex, $sText, $iSubItem, $iImage = -1)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iBuffer = StringLen($sText) + 1
Local $tBuffer
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tItem = DllStructCreate($tagLVITEM)
Local $iMask = $LVIF_TEXT
If $iImage <> -1 Then $iMask = BitOR($iMask, $LVIF_IMAGE)
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tItem, "Mask", $iMask)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "Image", $iImage)
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
DllStructSetData($tItem, "Text", $pBuffer)
$iRet = _SendMessage($hWnd, $LVM_SETITEMW, 0, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem + $iBuffer, $tMemMap)
Local $pText = $pMemory + $iItem
DllStructSetData($tItem, "Text", $pText)
_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $LVM_SETITEMW, 0, $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $LVM_SETITEMA, 0, $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
DllStructSetData($tItem, "Text", $pBuffer)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETITEMW, 0, $pItem)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETITEMA, 0, $pItem)
EndIf
EndIf
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_BeginUpdate($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $__LISTVIEWCONSTANT_WM_SETREDRAW, False) = 0
EndFunc
Func _GUICtrlListView_DeleteAllItems($hWnd)
If _GUICtrlListView_GetItemCount($hWnd) = 0 Then Return True
Local $vCID = 0
If IsHWnd($hWnd) Then
$vCID = _WinAPI_GetDlgCtrlID($hWnd)
Else
$vCID = $hWnd
$hWnd = GUICtrlGetHandle($hWnd)
EndIf
If $vCID < $_UDF_STARTID Then
Local $iParam = 0
For $iIndex = _GUICtrlListView_GetItemCount($hWnd) - 1 To 0 Step -1
$iParam = _GUICtrlListView_GetItemParam($hWnd, $iIndex)
If GUICtrlGetState($iParam) > 0 And GUICtrlGetHandle($iParam) = 0 Then
GUICtrlDelete($iParam)
EndIf
Next
If _GUICtrlListView_GetItemCount($hWnd) = 0 Then Return True
EndIf
Return _SendMessage($hWnd, $LVM_DELETEALLITEMS) <> 0
EndFunc
Func _GUICtrlListView_EndUpdate($hWnd)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Return _SendMessage($hWnd, $__LISTVIEWCONSTANT_WM_SETREDRAW, True) = 0
EndFunc
Func _GUICtrlListView_GetColumnCount($hWnd)
Return _SendMessage(_GUICtrlListView_GetHeader($hWnd), 0x1200)
EndFunc
Func _GUICtrlListView_GetHeader($hWnd)
If IsHWnd($hWnd) Then
Return HWnd(_SendMessage($hWnd, $LVM_GETHEADER))
Else
Return HWnd(GUICtrlSendMsg($hWnd, $LVM_GETHEADER, 0, 0))
EndIf
EndFunc
Func _GUICtrlListView_GetItemCount($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETITEMCOUNT)
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETITEMCOUNT, 0, 0)
EndIf
EndFunc
Func _GUICtrlListView_GetItemEx($hWnd, ByRef $tItem)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
$iRet = _SendMessage($hWnd, $LVM_GETITEMW, 0, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem, $tMemMap)
_MemWrite($tMemMap, $tItem)
If $bUnicode Then
_SendMessage($hWnd, $LVM_GETITEMW, 0, $pMemory, 0, "wparam", "ptr")
Else
_SendMessage($hWnd, $LVM_GETITEMA, 0, $pMemory, 0, "wparam", "ptr")
EndIf
_MemRead($tMemMap, $pMemory, $tItem, $iItem)
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_GETITEMW, 0, $pItem)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_GETITEMA, 0, $pItem)
EndIf
EndIf
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_GetItemParam($hWnd, $iIndex)
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tItem, "Mask", $LVIF_PARAM)
DllStructSetData($tItem, "Item", $iIndex)
_GUICtrlListView_GetItemEx($hWnd, $tItem)
Return DllStructGetData($tItem, "Param")
EndFunc
Func _GUICtrlListView_GetItemRect($hWnd, $iIndex, $iPart = 3)
Local $tRECT = _GUICtrlListView_GetItemRectEx($hWnd, $iIndex, $iPart)
Local $aRect[4]
$aRect[0] = DllStructGetData($tRECT, "Left")
$aRect[1] = DllStructGetData($tRECT, "Top")
$aRect[2] = DllStructGetData($tRECT, "Right")
$aRect[3] = DllStructGetData($tRECT, "Bottom")
Return $aRect
EndFunc
Func _GUICtrlListView_GetItemRectEx($hWnd, $iIndex, $iPart = 3)
Local $tRECT = DllStructCreate($tagRECT)
DllStructSetData($tRECT, "Left", $iPart)
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
_SendMessage($hWnd, $LVM_GETITEMRECT, $iIndex, $tRECT, 0, "wparam", "struct*")
Else
Local $iRect = DllStructGetSize($tRECT)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iRect, $tMemMap)
_MemWrite($tMemMap, $tRECT, $pMemory, $iRect)
_SendMessage($hWnd, $LVM_GETITEMRECT, $iIndex, $pMemory, 0, "wparam", "ptr")
_MemRead($tMemMap, $pMemory, $tRECT, $iRect)
_MemFree($tMemMap)
EndIf
Else
GUICtrlSendMsg($hWnd, $LVM_GETITEMRECT, $iIndex, DllStructGetPtr($tRECT))
EndIf
Return $tRECT
EndFunc
Func _GUICtrlListView_GetItemText($hWnd, $iIndex, $iSubItem = 0)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $tBuffer
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[4096]")
Else
$tBuffer = DllStructCreate("char Text[4096]")
EndIf
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "TextMax", 4096)
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
DllStructSetData($tItem, "Text", $pBuffer)
_SendMessage($hWnd, $LVM_GETITEMTEXTW, $iIndex, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem + 4096, $tMemMap)
Local $pText = $pMemory + $iItem
DllStructSetData($tItem, "Text", $pText)
_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
If $bUnicode Then
_SendMessage($hWnd, $LVM_GETITEMTEXTW, $iIndex, $pMemory, 0, "wparam", "ptr")
Else
_SendMessage($hWnd, $LVM_GETITEMTEXTA, $iIndex, $pMemory, 0, "wparam", "ptr")
EndIf
_MemRead($tMemMap, $pText, $tBuffer, 4096)
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
DllStructSetData($tItem, "Text", $pBuffer)
If $bUnicode Then
GUICtrlSendMsg($hWnd, $LVM_GETITEMTEXTW, $iIndex, $pItem)
Else
GUICtrlSendMsg($hWnd, $LVM_GETITEMTEXTA, $iIndex, $pItem)
EndIf
EndIf
Return DllStructGetData($tBuffer, "Text")
EndFunc
Func _GUICtrlListView_GetItemTextArray($hWnd, $iItem = -1)
Local $sItems = _GUICtrlListView_GetItemTextString($hWnd, $iItem)
If $sItems = "" Then
Local $aItems[1] = [0]
Return SetError($LV_ERR, $LV_ERR, $aItems)
EndIf
Return StringSplit($sItems, Opt('GUIDataSeparatorChar'))
EndFunc
Func _GUICtrlListView_GetItemTextString($hWnd, $iItem = -1)
Local $sRow = "", $sSeparatorChar = Opt('GUIDataSeparatorChar'), $iSelected
If $iItem = -1 Then
$iSelected = _GUICtrlListView_GetNextItem($hWnd)
Else
$iSelected = $iItem
EndIf
For $x = 0 To _GUICtrlListView_GetColumnCount($hWnd) - 1
$sRow &= _GUICtrlListView_GetItemText($hWnd, $iSelected, $x) & $sSeparatorChar
Next
Return StringTrimRight($sRow, 1)
EndFunc
Func _GUICtrlListView_GetNextItem($hWnd, $iStart = -1, $iSearch = 0, $iState = 8)
Local $aSearch[5] = [$LVNI_ALL, $LVNI_ABOVE, $LVNI_BELOW, $LVNI_TOLEFT, $LVNI_TORIGHT]
Local $iFlags = $aSearch[$iSearch]
If BitAND($iState, 1) <> 0 Then $iFlags = BitOR($iFlags, $LVNI_CUT)
If BitAND($iState, 2) <> 0 Then $iFlags = BitOR($iFlags, $LVNI_DROPHILITED)
If BitAND($iState, 4) <> 0 Then $iFlags = BitOR($iFlags, $LVNI_FOCUSED)
If BitAND($iState, 8) <> 0 Then $iFlags = BitOR($iFlags, $LVNI_SELECTED)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETNEXTITEM, $iStart, $iFlags)
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETNEXTITEM, $iStart, $iFlags)
EndIf
EndFunc
Func _GUICtrlListView_GetSelectedCount($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETSELECTEDCOUNT)
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETSELECTEDCOUNT, 0, 0)
EndIf
EndFunc
Func _GUICtrlListView_GetSelectedIndices($hWnd, $bArray = False)
Local $sIndices, $aIndices[1] = [0]
Local $iRet, $iCount = _GUICtrlListView_GetItemCount($hWnd)
For $iItem = 0 To $iCount
If IsHWnd($hWnd) Then
$iRet = _SendMessage($hWnd, $LVM_GETITEMSTATE, $iItem, $LVIS_SELECTED)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_GETITEMSTATE, $iItem, $LVIS_SELECTED)
EndIf
If $iRet Then
If(Not $bArray) Then
If StringLen($sIndices) Then
$sIndices &= "|" & $iItem
Else
$sIndices = $iItem
EndIf
Else
ReDim $aIndices[UBound($aIndices) + 1]
$aIndices[0] = UBound($aIndices) - 1
$aIndices[UBound($aIndices) - 1] = $iItem
EndIf
EndIf
Next
If(Not $bArray) Then
Return String($sIndices)
Else
Return $aIndices
EndIf
EndFunc
Func _GUICtrlListView_GetUnicodeFormat($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETUNICODEFORMAT) <> 0
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETUNICODEFORMAT, 0, 0) <> 0
EndIf
EndFunc
Func _GUICtrlListView_InsertColumn($hWnd, $iIndex, $sText, $iWidth = 50, $iAlign = -1, $iImage = -1, $bOnRight = False)
Local $aAlign[3] = [$LVCFMT_LEFT, $LVCFMT_RIGHT, $LVCFMT_CENTER]
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iBuffer = StringLen($sText) + 1
Local $tBuffer
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tColumn = DllStructCreate($tagLVCOLUMN)
Local $iMask = BitOR($LVCF_FMT, $LVCF_WIDTH, $LVCF_TEXT)
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
Local $iFmt = $aAlign[$iAlign]
If $iImage <> -1 Then
$iMask = BitOR($iMask, $LVCF_IMAGE)
$iFmt = BitOR($iFmt, $LVCFMT_COL_HAS_IMAGES, $LVCFMT_IMAGE)
EndIf
If $bOnRight Then $iFmt = BitOR($iFmt, $LVCFMT_BITMAP_ON_RIGHT)
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tColumn, "Mask", $iMask)
DllStructSetData($tColumn, "Fmt", $iFmt)
DllStructSetData($tColumn, "CX", $iWidth)
DllStructSetData($tColumn, "TextMax", $iBuffer)
DllStructSetData($tColumn, "Image", $iImage)
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
DllStructSetData($tColumn, "Text", $pBuffer)
$iRet = _SendMessage($hWnd, $LVM_INSERTCOLUMNW, $iIndex, $tColumn, 0, "wparam", "struct*")
Else
Local $iColumn = DllStructGetSize($tColumn)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iColumn + $iBuffer, $tMemMap)
Local $pText = $pMemory + $iColumn
DllStructSetData($tColumn, "Text", $pText)
_MemWrite($tMemMap, $tColumn, $pMemory, $iColumn)
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $LVM_INSERTCOLUMNW, $iIndex, $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $LVM_INSERTCOLUMNA, $iIndex, $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Else
Local $pColumn = DllStructGetPtr($tColumn)
DllStructSetData($tColumn, "Text", $pBuffer)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_INSERTCOLUMNW, $iIndex, $pColumn)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_INSERTCOLUMNA, $iIndex, $pColumn)
EndIf
EndIf
If $iAlign > 0 Then _GUICtrlListView_SetColumn($hWnd, $iRet, $sText, $iWidth, $iAlign, $iImage, $bOnRight)
Return $iRet
EndFunc
Func _GUICtrlListView_InsertItem($hWnd, $sText, $iIndex = -1, $iImage = -1, $iParam = 0)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $iBuffer, $tBuffer, $iRet
If $iIndex = -1 Then $iIndex = 999999999
Local $tItem = DllStructCreate($tagLVITEM)
DllStructSetData($tItem, "Param", $iParam)
$iBuffer = StringLen($sText) + 1
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tItem, "Text", DllStructGetPtr($tBuffer))
DllStructSetData($tItem, "TextMax", $iBuffer)
Local $iMask = BitOR($LVIF_TEXT, $LVIF_PARAM)
If $iImage >= 0 Then $iMask = BitOR($iMask, $LVIF_IMAGE)
DllStructSetData($tItem, "Mask", $iMask)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "Image", $iImage)
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Or($sText = -1) Then
$iRet = _SendMessage($hWnd, $LVM_INSERTITEMW, 0, $tItem, 0, "wparam", "struct*")
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iItem + $iBuffer, $tMemMap)
Local $pText = $pMemory + $iItem
DllStructSetData($tItem, "Text", $pText)
_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $LVM_INSERTITEMW, 0, $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $LVM_INSERTITEMA, 0, $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Else
Local $pItem = DllStructGetPtr($tItem)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_INSERTITEMW, 0, $pItem)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_INSERTITEMA, 0, $pItem)
EndIf
EndIf
Return $iRet
EndFunc
Func _GUICtrlListView_SetColumn($hWnd, $iIndex, $sText, $iWidth = -1, $iAlign = -1, $iImage = -1, $bOnRight = False)
Local $bUnicode = _GUICtrlListView_GetUnicodeFormat($hWnd)
Local $aAlign[3] = [$LVCFMT_LEFT, $LVCFMT_RIGHT, $LVCFMT_CENTER]
Local $iBuffer = StringLen($sText) + 1
Local $tBuffer
If $bUnicode Then
$tBuffer = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tBuffer = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
Local $pBuffer = DllStructGetPtr($tBuffer)
Local $tColumn = DllStructCreate($tagLVCOLUMN)
Local $iMask = $LVCF_TEXT
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
$iMask = BitOR($iMask, $LVCF_FMT)
Local $iFmt = $aAlign[$iAlign]
If $iWidth <> -1 Then $iMask = BitOR($iMask, $LVCF_WIDTH)
If $iImage <> -1 Then
$iMask = BitOR($iMask, $LVCF_IMAGE)
$iFmt = BitOR($iFmt, $LVCFMT_COL_HAS_IMAGES, $LVCFMT_IMAGE)
Else
$iImage = 0
EndIf
If $bOnRight Then $iFmt = BitOR($iFmt, $LVCFMT_BITMAP_ON_RIGHT)
DllStructSetData($tBuffer, "Text", $sText)
DllStructSetData($tColumn, "Mask", $iMask)
DllStructSetData($tColumn, "Fmt", $iFmt)
DllStructSetData($tColumn, "CX", $iWidth)
DllStructSetData($tColumn, "TextMax", $iBuffer)
DllStructSetData($tColumn, "Image", $iImage)
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hLVLastWnd) Then
DllStructSetData($tColumn, "Text", $pBuffer)
$iRet = _SendMessage($hWnd, $LVM_SETCOLUMNW, $iIndex, $tColumn, 0, "wparam", "struct*")
Else
Local $iColumn = DllStructGetSize($tColumn)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iColumn + $iBuffer, $tMemMap)
Local $pText = $pMemory + $iColumn
DllStructSetData($tColumn, "Text", $pText)
_MemWrite($tMemMap, $tColumn, $pMemory, $iColumn)
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $LVM_SETCOLUMNW, $iIndex, $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $LVM_SETCOLUMNA, $iIndex, $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Else
Local $pColumn = DllStructGetPtr($tColumn)
DllStructSetData($tColumn, "Text", $pBuffer)
If $bUnicode Then
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETCOLUMNW, $iIndex, $pColumn)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETCOLUMNA, $iIndex, $pColumn)
EndIf
EndIf
Return $iRet <> 0
EndFunc
#Au3Stripper_Ignore_Funcs=__GUICtrlListView_Sort
Func __GUICtrlListView_Sort($nItem1, $nItem2, $hWnd)
Local $iIndex, $sVal1, $sVal2, $nResult
For $x = 1 To $__g_aListViewSortInfo[0][0]
If $hWnd = $__g_aListViewSortInfo[$x][1] Then
$iIndex = $x
ExitLoop
EndIf
Next
If $__g_aListViewSortInfo[$iIndex][3] = $__g_aListViewSortInfo[$iIndex][4] Then
If Not $__g_aListViewSortInfo[$iIndex][7] Then
$__g_aListViewSortInfo[$iIndex][5] *= -1
$__g_aListViewSortInfo[$iIndex][7] = 1
EndIf
Else
$__g_aListViewSortInfo[$iIndex][7] = 1
EndIf
$__g_aListViewSortInfo[$iIndex][6] = $__g_aListViewSortInfo[$iIndex][3]
$sVal1 = _GUICtrlListView_GetItemText($hWnd, $nItem1, $__g_aListViewSortInfo[$iIndex][3])
$sVal2 = _GUICtrlListView_GetItemText($hWnd, $nItem2, $__g_aListViewSortInfo[$iIndex][3])
If $__g_aListViewSortInfo[$iIndex][8] = 1 Then
If(StringIsFloat($sVal1) Or StringIsInt($sVal1)) Then $sVal1 = Number($sVal1)
If(StringIsFloat($sVal2) Or StringIsInt($sVal2)) Then $sVal2 = Number($sVal2)
EndIf
If $__g_aListViewSortInfo[$iIndex][8] < 2 Then
$nResult = 0
If $sVal1 < $sVal2 Then
$nResult = -1
ElseIf $sVal1 > $sVal2 Then
$nResult = 1
EndIf
Else
$nResult = DllCall('shlwapi.dll', 'int', 'StrCmpLogicalW', 'wstr', $sVal1, 'wstr', $sVal2)[0]
EndIf
$nResult = $nResult * $__g_aListViewSortInfo[$iIndex][5]
Return $nResult
EndFunc
Global Const $__STATUSBARCONSTANT_WM_USER = 0X400
Global Const $SB_GETBORDERS =($__STATUSBARCONSTANT_WM_USER + 7)
Global Const $SB_GETRECT =($__STATUSBARCONSTANT_WM_USER + 10)
Global Const $SB_GETUNICODEFORMAT = 0x2000 + 6
Global Const $SB_ISSIMPLE =($__STATUSBARCONSTANT_WM_USER + 14)
Global Const $SB_SETPARTS =($__STATUSBARCONSTANT_WM_USER + 4)
Global Const $SB_SETTEXTA =($__STATUSBARCONSTANT_WM_USER + 1)
Global Const $SB_SETTEXTW =($__STATUSBARCONSTANT_WM_USER + 11)
Global Const $SB_SETTEXT = $SB_SETTEXTA
Global Const $SB_SIMPLEID = 0xff
Global $__g_hSBLastWnd
Global Const $__STATUSBARCONSTANT_ClassName = "msctls_statusbar32"
Global Const $__STATUSBARCONSTANT_WM_SIZE = 0x05
Global Const $tagBORDERS = "int BX;int BY;int RX"
Func _GUICtrlStatusBar_Create($hWnd, $vPartEdge = -1, $vPartText = "", $iStyles = -1, $iExStyles = 0x00000000)
If Not IsHWnd($hWnd) Then Return SetError(1, 0, 0)
Local $iStyle = BitOR($__UDFGUICONSTANT_WS_CHILD, $__UDFGUICONSTANT_WS_VISIBLE)
If $iStyles = -1 Then $iStyles = 0x00000000
If $iExStyles = -1 Then $iExStyles = 0x00000000
Local $aPartWidth[1], $aPartText[1]
If @NumParams > 1 Then
If IsArray($vPartEdge) Then
$aPartWidth = $vPartEdge
Else
$aPartWidth[0] = $vPartEdge
EndIf
If @NumParams = 2 Then
ReDim $aPartText[UBound($aPartWidth)]
Else
If IsArray($vPartText) Then
$aPartText = $vPartText
Else
$aPartText[0] = $vPartText
EndIf
If UBound($aPartWidth) <> UBound($aPartText) Then
Local $iLast
If UBound($aPartWidth) > UBound($aPartText) Then
$iLast = UBound($aPartText)
ReDim $aPartText[UBound($aPartWidth)]
Else
$iLast = UBound($aPartWidth)
ReDim $aPartWidth[UBound($aPartText)]
For $x = $iLast To UBound($aPartWidth) - 1
$aPartWidth[$x] = $aPartWidth[$x - 1] + 75
Next
$aPartWidth[UBound($aPartText) - 1] = -1
EndIf
EndIf
EndIf
If Not IsHWnd($hWnd) Then $hWnd = HWnd($hWnd)
If @NumParams > 3 Then $iStyle = BitOR($iStyle, $iStyles)
EndIf
Local $nCtrlID = __UDF_GetNextGlobalID($hWnd)
If @error Then Return SetError(@error, @extended, 0)
Local $hWndSBar = _WinAPI_CreateWindowEx($iExStyles, $__STATUSBARCONSTANT_ClassName, "", $iStyle, 0, 0, 0, 0, $hWnd, $nCtrlID)
If @error Then Return SetError(@error, @extended, 0)
If @NumParams > 1 Then
_GUICtrlStatusBar_SetParts($hWndSBar, UBound($aPartWidth), $aPartWidth)
For $x = 0 To UBound($aPartText) - 1
_GUICtrlStatusBar_SetText($hWndSBar, $aPartText[$x], $x)
Next
EndIf
Return $hWndSBar
EndFunc
Func _GUICtrlStatusBar_EmbedControl($hWnd, $iPart, $hControl, $iFit = 4)
Local $aRect = _GUICtrlStatusBar_GetRect($hWnd, $iPart)
Local $iBarX = $aRect[0]
Local $iBarY = $aRect[1]
Local $iBarW = $aRect[2] - $iBarX
Local $iBarH = $aRect[3] - $iBarY
Local $iConX = $iBarX
Local $iConY = $iBarY
Local $iConW = _WinAPI_GetWindowWidth($hControl)
Local $iConH = _WinAPI_GetWindowHeight($hControl)
If $iConW > $iBarW Then $iConW = $iBarW
If $iConH > $iBarH Then $iConH = $iBarH
Local $iPadX =($iBarW - $iConW) / 2
Local $iPadY =($iBarH - $iConH) / 2
If $iPadX < 0 Then $iPadX = 0
If $iPadY < 0 Then $iPadY = 0
If BitAND($iFit, 1) = 1 Then $iConX = $iBarX + $iPadX
If BitAND($iFit, 2) = 2 Then $iConY = $iBarY + $iPadY
If BitAND($iFit, 4) = 4 Then
$iPadX = _GUICtrlStatusBar_GetBordersRect($hWnd)
$iPadY = _GUICtrlStatusBar_GetBordersVert($hWnd)
$iConX = $iBarX
If _GUICtrlStatusBar_IsSimple($hWnd) Then $iConX += $iPadX
$iConY = $iBarY + $iPadY
$iConW = $iBarW -($iPadX * 2)
$iConH = $iBarH -($iPadY * 2)
EndIf
_WinAPI_SetParent($hControl, $hWnd)
_WinAPI_MoveWindow($hControl, $iConX, $iConY, $iConW, $iConH)
EndFunc
Func _GUICtrlStatusBar_GetBorders($hWnd)
Local $tBorders = DllStructCreate($tagBORDERS)
Local $iRet
If _WinAPI_InProcess($hWnd, $__g_hSBLastWnd) Then
$iRet = _SendMessage($hWnd, $SB_GETBORDERS, 0, $tBorders, 0, "wparam", "struct*")
Else
Local $iSize = DllStructGetSize($tBorders)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iSize, $tMemMap)
$iRet = _SendMessage($hWnd, $SB_GETBORDERS, 0, $pMemory, 0, "wparam", "ptr")
_MemRead($tMemMap, $pMemory, $tBorders, $iSize)
_MemFree($tMemMap)
EndIf
Local $aBorders[3]
If $iRet = 0 Then Return SetError(-1, -1, $aBorders)
$aBorders[0] = DllStructGetData($tBorders, "BX")
$aBorders[1] = DllStructGetData($tBorders, "BY")
$aBorders[2] = DllStructGetData($tBorders, "RX")
Return $aBorders
EndFunc
Func _GUICtrlStatusBar_GetBordersRect($hWnd)
Local $aBorders = _GUICtrlStatusBar_GetBorders($hWnd)
Return SetError(@error, @extended, $aBorders[2])
EndFunc
Func _GUICtrlStatusBar_GetBordersVert($hWnd)
Local $aBorders = _GUICtrlStatusBar_GetBorders($hWnd)
Return SetError(@error, @extended, $aBorders[1])
EndFunc
Func _GUICtrlStatusBar_GetRect($hWnd, $iPart)
Local $tRECT = _GUICtrlStatusBar_GetRectEx($hWnd, $iPart)
If @error Then Return SetError(@error, 0, 0)
Local $aRect[4]
$aRect[0] = DllStructGetData($tRECT, "Left")
$aRect[1] = DllStructGetData($tRECT, "Top")
$aRect[2] = DllStructGetData($tRECT, "Right")
$aRect[3] = DllStructGetData($tRECT, "Bottom")
Return $aRect
EndFunc
Func _GUICtrlStatusBar_GetRectEx($hWnd, $iPart)
Local $tRECT = DllStructCreate($tagRECT)
Local $iRet
If _WinAPI_InProcess($hWnd, $__g_hSBLastWnd) Then
$iRet = _SendMessage($hWnd, $SB_GETRECT, $iPart, $tRECT, 0, "wparam", "struct*")
Else
Local $iRect = DllStructGetSize($tRECT)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iRect, $tMemMap)
$iRet = _SendMessage($hWnd, $SB_GETRECT, $iPart, $pMemory, 0, "wparam", "ptr")
_MemRead($tMemMap, $pMemory, $tRECT, $iRect)
_MemFree($tMemMap)
EndIf
Return SetError($iRet = 0, 0, $tRECT)
EndFunc
Func _GUICtrlStatusBar_GetUnicodeFormat($hWnd)
Return _SendMessage($hWnd, $SB_GETUNICODEFORMAT) <> 0
EndFunc
Func _GUICtrlStatusBar_IsSimple($hWnd)
Return _SendMessage($hWnd, $SB_ISSIMPLE) <> 0
EndFunc
Func _GUICtrlStatusBar_Resize($hWnd)
_SendMessage($hWnd, $__STATUSBARCONSTANT_WM_SIZE)
EndFunc
Func _GUICtrlStatusBar_SetParts($hWnd, $vPartEdge = -1, $vPartWidth = 25)
If IsArray($vPartEdge) And IsArray($vPartWidth) Then Return False
Local $tParts, $iParts
If IsArray($vPartEdge) Then
$vPartEdge[UBound($vPartEdge) - 1] = -1
$iParts = UBound($vPartEdge)
$tParts = DllStructCreate("int[" & $iParts & "]")
For $x = 0 To $iParts - 2
DllStructSetData($tParts, 1, $vPartEdge[$x], $x + 1)
Next
DllStructSetData($tParts, 1, -1, $iParts)
Else
If $vPartEdge < -1 Then Return False
If IsArray($vPartWidth) Then
$iParts = UBound($vPartWidth)
$tParts = DllStructCreate("int[" & $iParts & "]")
Local $iPartRightEdge = 0
For $x = 0 To $iParts - 2
$iPartRightEdge += $vPartWidth[$x]
If $vPartWidth[$x] <= 0 Then Return False
DllStructSetData($tParts, 1, $iPartRightEdge, $x + 1)
Next
DllStructSetData($tParts, 1, -1, $iParts)
ElseIf $vPartEdge > 1 Then
$iParts = $vPartEdge
$tParts = DllStructCreate("int[" & $iParts & "]")
For $x = 1 To $iParts - 1
DllStructSetData($tParts, 1, $vPartWidth * $x, $x)
Next
DllStructSetData($tParts, 1, -1, $iParts)
Else
$iParts = 1
$tParts = DllStructCreate("int")
DllStructSetData($tParts, 1, -1)
EndIf
EndIf
If _WinAPI_InProcess($hWnd, $__g_hSBLastWnd) Then
_SendMessage($hWnd, $SB_SETPARTS, $iParts, $tParts, 0, "wparam", "struct*")
Else
Local $iSize = DllStructGetSize($tParts)
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iSize, $tMemMap)
_MemWrite($tMemMap, $tParts)
_SendMessage($hWnd, $SB_SETPARTS, $iParts, $pMemory, 0, "wparam", "ptr")
_MemFree($tMemMap)
EndIf
_GUICtrlStatusBar_Resize($hWnd)
Return True
EndFunc
Func _GUICtrlStatusBar_SetText($hWnd, $sText = "", $iPart = 0, $iUFlag = 0)
Local $bUnicode = _GUICtrlStatusBar_GetUnicodeFormat($hWnd)
Local $iBuffer = StringLen($sText) + 1
Local $tText
If $bUnicode Then
$tText = DllStructCreate("wchar Text[" & $iBuffer & "]")
$iBuffer *= 2
Else
$tText = DllStructCreate("char Text[" & $iBuffer & "]")
EndIf
DllStructSetData($tText, "Text", $sText)
If _GUICtrlStatusBar_IsSimple($hWnd) Then $iPart = $SB_SIMPLEID
Local $iRet
If _WinAPI_InProcess($hWnd, $__g_hSBLastWnd) Then
$iRet = _SendMessage($hWnd, $SB_SETTEXTW, BitOR($iPart, $iUFlag), $tText, 0, "wparam", "struct*")
Else
Local $tMemMap
Local $pMemory = _MemInit($hWnd, $iBuffer, $tMemMap)
_MemWrite($tMemMap, $tText)
If $bUnicode Then
$iRet = _SendMessage($hWnd, $SB_SETTEXTW, BitOR($iPart, $iUFlag), $pMemory, 0, "wparam", "ptr")
Else
$iRet = _SendMessage($hWnd, $SB_SETTEXT, BitOR($iPart, $iUFlag), $pMemory, 0, "wparam", "ptr")
EndIf
_MemFree($tMemMap)
EndIf
Return $iRet <> 0
EndFunc
Global Const $PBS_SMOOTH = 1
Global $g_sBotFile = "multibot.run.exe"
Global $g_sBotFileAU3 = "multibot.run.au3"
Global $g_sVersion = "1.0.4"
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
Case $hMenu_MEmu
ShellExecute("https://forum.multibot.run/index.php?resources/multibotrun-memu-v5-2-3.61/")
Case $hMenu_Nox
ShellExecute("https://forum.multibot.run/index.php?resources/multibotrun-nox-7-0-1-5.62/")
Case $hMenu_CheckForUpdate
$sTempPath = _WinAPI_GetTempFileName(@TempDir)
$hUpdateFile = InetGet("https://raw.githubusercontent.com/promac2k/SelectMultiBotRun/master/SelectMultiBotRun_Info.txt", $sTempPath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
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
EndFunc
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
EndFunc
Func GUI_Emulator()
$g_hGui_Emulator = GUICreate("Emulator", 258, 167, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
$hCmb_Emulator = GUICtrlCreateCombo("MEmu", 24, 72, 201, 21, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "MEmu|Nox")
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
EndFunc
Func GUI_Instance()
Local $hLbl_Instance = 0
$g_hGui_Instance = GUICreate("Instance", 258, 167, $g_aGuiPos_Main[0], $g_aGuiPos_Main[1] + 150, -1, -1, $g_hGui_Main)
$hIpt_Instance = GUICtrlCreateInput("", 24, 72, 201, 21)
$hBtn_Next = GUICtrlCreateButton("Next step", 72, 120, 97, 25, $WS_GROUP)
$hLbl_Instance = GUICtrlCreateLabel("Please type in the Instance Name you want to use", 24, 8, 204, 57)
GUISetState(@SW_HIDE, $g_hGui_Instance)
Switch $g_sSelectedEmulator
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
Case "BlueStacks3"
$Instance = StringRegExp($Instances, "(?i)" & "Android" & "(?:[_][0-9])?", 3)
Case "iTools"
$Instance = StringRegExp($Instances, "(?)iToolsVM(?:[_][0-9][0-9])?", 3)
Case "LeapDroid"
$Instance = StringRegExp($Instances, "(?i)vm\d?", 3)
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
EndFunc
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
EndFunc
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
EndFunc
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
Case "MEmu"
GUICtrlSetData($hCmb_Emulator, "Nox")
Case "Nox"
GUICtrlSetData($hCmb_Emulator, "MEmu")
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
If $sSelectedEmulator = "BlueStacks" Or $sSelectedEmulator = "BlueStacks2" Then
GUICtrlSetState($hIpt_Instance, $GUI_DISABLE)
GUICtrlSetData($hIpt_Instance, "")
ElseIf $sSelectedEmulator <> "BlueStacks" And "BlueStacks2" Then
GUICtrlSetState($hIpt_Instance, $GUI_ENABLE)
Switch $sSelectedEmulator
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
EndFunc
Func GUI_AutoStart()
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
FileCreateShortcut($g_sIniDir & "\" & $g_sBotFile, @StartupDir & "\MultiBot -" & $g_sIniProfile & ".lnk", $g_sIniDir, $g_sIniProfile & " " & $g_sIniEmulator = "BlueStacks3" ? "BlueStacks2" : $g_sIniEmulator & " " & $g_sIniInstance, "Shortcut for Bot Profile:" & $g_sIniProfile)
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
EndFunc
Func RunSetup()
$Lstbx_Sel = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
If $Lstbx_Sel[0] > 0 Then
For $i = 1 To $Lstbx_Sel[0]
$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $Lstbx_Sel[$i])
If $sLstbx_SelItem <> "" Then
ReadIni($sLstbx_SelItem)
Local $sEmulator = $g_sIniEmulator
If $g_sIniEmulator = "BlueStacks3" Then $sEmulator = "BlueStacks2"
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
EndFunc
Func CreateShortcut()
Local $iCreatedSC = 0, $sBotFileName, $hSC
$Lstbx_Sel = _GUICtrlListView_GetSelectedIndices($g_hListview_Main, True)
If $Lstbx_Sel[0] > 0 Then
For $i = 1 To $Lstbx_Sel[0]
$sLstbx_SelItem = _GUICtrlListView_GetItemText($g_hListview_Main, $Lstbx_Sel[$i])
If $sLstbx_SelItem <> "" Then
ReadIni($sLstbx_SelItem)
Local $sEmulator = $g_sIniEmulator
If $g_sIniEmulator = "BlueStacks3" Then $sEmulator = "BlueStacks2"
$aParameters = StringSplit($g_sIniParameters, "")
Local $sSpecialParameter = $aParameters[1] = 1 ? " /nowatchdog" : "" & $aParameters[2] = 1 ? " /dock1" : "" & $aParameters[3] = 1 ? " /dock2" : "" & $aParameters[4] = 1 ? " /dpiaware" : "" & $aParameters[5] = 1 ? " /debug" : "" & $aParameters[6] = 1 ? " /minigui" : "" & $aParameters[7] = 1 ? " /hideandroid" : ""
If FileExists($g_sIniDir & "\" & $g_sBotFile) Then
$sBotFileName = $g_sBotFile
ElseIf FileExists($g_sIniDir & "\" & $g_sBotFileAU3) Then
$sBotFileName = $g_sBotFileAU3
Else
MsgBox($MB_OK, "No Bot found", "Couldn't find any Bot in the Directory, please check if you have the multibot.run.exe or the multibot.run.au3 in the Dir and if you selected the right Dir!", 0, $g_hGui_Main)
EndIf
$hSC = FileCreateShortcut($g_sIniDir & "\" & $sBotFileName, @DesktopDir & "\MultiBot -" & $g_sIniProfile & ".lnk", $g_sIniDir, $g_sIniProfile & " " & $sEmulator & " " & $g_sIniInstance & $sSpecialParameter, "Shortcut for Bot Profile:" & $g_sIniProfile)
If $hSC = 1 Then $iCreatedSC += 1
EndIf
Next
If $iCreatedSC = 1 Then
_GUICtrlStatusBar_SetText($g_hLog, "Created " & $iCreatedSC & " Shortcut")
ElseIf $iCreatedSC > 1 Then
_GUICtrlStatusBar_SetText($g_hLog, "Created " & $iCreatedSC & " Shortcuts")
EndIf
$iCreatedSC = 0
EndIf
EndFunc
Func UpdateList_Main()
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
EndFunc
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
EndFunc
Func GetBotVers()
Local $aSections
$aSections = IniReadSectionNames($g_sDirProfiles)
For $i = 1 To UBound($aSections, 1) - 1
ReadIni($aSections[$i])
If $aSections[$i] = "Options" Then ContinueLoop
$hBotVers = FileOpen($g_sIniDir & "\multibot.run.version.au3")
$sBotVers = FileRead($hBotVers)
$aBotVers = StringRegExp($sBotVers, "(?i)v([0-9]{1,3}).([0-9]{1,3}).([0-9]{1,3})", 2)
FileClose($hBotVers)
If IsArray($aBotVers) Then
IniWrite($g_sDirProfiles, $aSections[$i], "BotVers", $aBotVers[0])
EndIf
Next
EndFunc
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
EndFunc
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
If($iY >= $aRect[1]) And($iY <= $aRect[3]) Then _ContextMenu($i)
Next
Return $GUI_RUNDEFMSG
ElseIf _GUICtrlListView_GetItemCount($g_hListview_Main) < 1 Then
Return
EndIf
EndFunc
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
EndFunc
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
EndFunc
Func UpdateSelect()
FileMove(@ScriptDir & "\" & @ScriptName, @ScriptDir & "\" & "SelectMultiBotRunOLD" & $g_sVersion & ".exe")
$hUpdateFile = InetGet("https://raw.githubusercontent.com/promac2k/SelectMultiBotRun/master/SelectMultiBotRun.exe", @ScriptDir & "\SelectMultiBotRun.exe", $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
Do
Sleep(250)
Until InetGetInfo($hUpdateFile, $INET_DOWNLOADCOMPLETE)
InetClose($hUpdateFile)
MsgBox($MB_OK, "Update", "Update finished! Downloaded newer Version and placed it in old Directoy!.New Version gets now started! Just delete the SelectMultiBotRunOLD" & $g_sVersion & ".exe and continue botting :)", 0, $g_hGui_Main)
ShellExecute(@ScriptDir & "\SelectMultiBotRun.exe")
Exit
EndFunc
Func CheckUpdate()
$sTempPath = @MyDocumentsDir & "\SelectMultiBotRun_Info.txt"
$hUpdateFile = InetGet("https://raw.githubusercontent.com/promac2k/SelectMultiBotRun/master/SelectMultiBotRun_Info.txt", $sTempPath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
Do
Sleep(250)
Until InetGetInfo($hUpdateFile, $INET_DOWNLOADCOMPLETE)
InetClose($hUpdateFile)
$hGitVersion = IniRead($sTempPath, "General", "DisplayVers", "")
$sGitVersion = StringStripWS($hGitVersion, 8)
$Update = _VersionCompare($g_sVersion, $sGitVersion)
If $Update = -1 Then GUICtrlSetState($g_hLblUpdateAvailable, $GUI_SHOW)
FileDelete($sTempPath)
EndFunc
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
EndFunc
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
EndFunc
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
EndFunc
Func GetNoxPath()
Local $sNoxPath = RegRead($HKLM & "\SOFTWARE" & $Wow6432Node & "\DuoDianOnline\SetupInfo\", "InstallPath")
If @error = 0 Then
If StringRight($sNoxPath, 1) <> "\" Then $sNoxPath &= "\"
$sNoxPath &= "bin\"
Else
$sNoxPath = ""
EndIf
Return $sNoxPath
EndFunc
Func GetNoxRtPath()
Local $sNoxRtPath = RegRead($HKLM & "\SOFTWARE\BigNox\VirtualBox\", "InstallDir")
If @error = 0 Then
If StringRight($sNoxRtPath, 1) <> "\" Then $sNoxRtPath &= "\"
Else
$sNoxRtPath = @ProgramFilesDir & "\Bignox\BigNoxVM\RT\"
EndIf
Return $sNoxRtPath
EndFunc
Func IsAndroidInstalled($sAndroid)
Local $sPath, $sFile, $bIsInstalled = False
Switch $sAndroid
Case "MEmu"
$sPath = GetMEmuPath()
$sFile = "MEmu.exe"
Case "Nox"
$sPath = GetNoxPath()
$sFile = "Nox.exe"
EndSwitch
If FileExists($sPath & $sFile) = 1 Then $bIsInstalled = True
Return $bIsInstalled
EndFunc
Func GetInstanceMgrPath($sAndroid)
Local $sManagerPath
Switch $sAndroid
Case "MEmu"
$sManagerPath = EnvGet("MEmuHyperv_Path") & "\MEmuManage.exe"
If FileExists($sManagerPath) = 0 Then
$sManagerPath = GetMEmuPath() & "..\MEmuHyperv\MEmuManage.exe"
EndIf
Case "Nox"
$sManagerPath = GetNoxRtPath() & "BigNoxVMMgr.exe"
EndSwitch
Return $sManagerPath
EndFunc
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
If($iTimeOut > 0 And TimerDiff($hTimer) > $iTimeOut) Then ExitLoop
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
EndFunc
Func CleanLaunchOutput(ByRef $output)
$output = StringReplace($output, @CR & @CR, "")
$output = StringReplace($output, @CRLF & @CRLF, "")
If StringRight($output, 1) = @LF Then $output = StringLeft($output, StringLen($output) - 1)
If StringRight($output, 1) = @CR Then $output = StringLeft($output, StringLen($output) - 1)
EndFunc
