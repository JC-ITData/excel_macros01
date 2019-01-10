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
