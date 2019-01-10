Option Explicit
' contains functions you use in cell formulas
Function ConcatIf(CriteriaRange As Range, strMatchCriteria As Integer, ConcatRange As Range, Seperator As String)
' purpose: similar to SUMIF, except its used to concatenate
' parameters:
	' CriteriaRange: cells to evaluate
    ' strMatchCriteria
	' ConcatRange: cells to concatenate

    Dim c As Range
    Dim strResults As String
    Dim strFirstAddress As String
    Dim intOffset As Integer
    
    ' assumes ConcatRange is same shape as CriteriaRange and on same rows. i.e. ConcatRange is a column offset of CriteriaRange.
    ' <- later - add some checks / error handling
    intOffset = ConcatRange.Column - CriteriaRange.Column
    
    strResults = ""
    
    'For Each c In CriteriaRange.Rows("2:" & CriteriaRange.Rows.Count)
    '    MsgBox "Address: " + c.Address
    'Next c
    
    Set c = CriteriaRange.Find(What:=strMatchCriteria, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False, after:=CriteriaRange.Cells(1, 1))
                
    If Not (c Is Nothing) Then
        strFirstAddress = c.Address
        
        Do
            strResults = strResults + c.Offset(0, intOffset).Value + Seperator
            'Set c = CriteriaRange.FindNext(after:=c)
            Set c = CriteriaRange.Find(What:=strMatchCriteria, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False, after:=c)
        Loop While Not c Is Nothing And c.Address <> strFirstAddress
        
        strResults = Left(strResults, Len(strResults) - Len(Seperator))
        
    End If
    
    ConcatIf = strResults
    
End Function

Function GetColAlpha(intCol As Integer) As String
    Dim intNumAs As Integer
    
    'intNumAs = Int(intCol / 26) ' This logic does not work for cols ending in Z, so replaced with IF statement below.
    If (intCol / 26) - Int(intCol / 26) > 0 Then
        intNumAs = Int(intCol / 26)
    Else
        intNumAs = Int((intCol - 1) / 26)
    End If
    If intNumAs >= 1 Then GetColAlpha = Chr(64 + intNumAs)
    GetColAlpha = GetColAlpha & Chr(intCol - (intNumAs * 26) + 64)
    
End Function
Function GetColInt(strCol As String) As Integer
    Dim intCol As Integer
    Dim intTemp As Integer
    Dim i As Integer
    Dim intLen As Integer
    
    intLen = Len(strCol)
    intCol = 0
    If intLen > 1 Then
        intCol = (Asc(Left(strCol, 1)) - 64) * 26
    End If
    intCol = intCol + (Asc(Right(strCol, 1)) - 64)
    GetColInt = intCol
End Function
