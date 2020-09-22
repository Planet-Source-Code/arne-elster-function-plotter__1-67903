Attribute VB_Name = "modFormula"
Option Explicit

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" ( _
    ptr() As Any _
) As Long

Private Declare Sub CopySafeArray Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal dst As Long, src As Any, ByVal nBytes As Long _
)

Private Declare Sub DestroySafeArray Lib "kernel32" Alias "RtlZeroMemory" ( _
    ByVal dst As Long, ByVal nBytes As Long _
)

Private Type SAFEARRAY1D
    cDims                   As Integer
    fFeatures               As Integer
    cbElements              As Long
    cLocks                  As Long
    pvData                  As Long
    cElements               As Long
    lLbound                 As Long
End Type

Private Type TFunction
    name                    As String
    params                  As Long
End Type

Private Type Variable
    DValue                  As Double
    VarName                 As String
End Type

Public Type Token
    ttype                   As Tokens
    TValue                  As String
    TDValue                 As Double
    SyIdx                   As Long
    OpPrec                  As Long
End Type

Private Const CHAR_PLUS     As Long = 43
Private Const CHAR_MINUS    As Long = 45
Private Const CHAR_ASTERISK As Long = 42
Private Const CHAR_SLASH    As Long = 47
Private Const CHAR_POWER    As Long = 94
Private Const CHAR_PARL     As Long = 40
Private Const CHAR_PARR     As Long = 41
Private Const CHAR_SEP      As Long = 44
Private Const CHAR_WS       As Long = 32

Private Const TOKEN_OP_FLAG As Long = &H80000000

Private Const OP_ANY_PREC   As Long = -1
Private Const OP_ADD_PREC   As Long = OP_ANY_PREC + 1
Private Const OP_SUB_PREC   As Long = OP_ADD_PREC
Private Const OP_MUL_PREC   As Long = OP_ADD_PREC + 1
Private Const OP_DIV_PREC   As Long = OP_MUL_PREC
Private Const OP_POW_PREC   As Long = OP_MUL_PREC + 1

Public Enum Tokens
    TokenNumber = &H1
    
    TokenAdd = &H2 Or TOKEN_OP_FLAG
    TokenSub = &H4 Or TOKEN_OP_FLAG
    TokenMul = &H8 Or TOKEN_OP_FLAG
    TokenDiv = &H10 Or TOKEN_OP_FLAG
    TokenPow = &H20 Or TOKEN_OP_FLAG
    
    TokenParL = &H40
    TokenParR = &H80
    TokenVar = &H100
    TokenFnc = &H200
    
    TokenSep = &H8000000
    TokenBegin = &H10000000
    TokenEnd = &H20000000
    
    TokenUnknown = &H40000000
End Enum

Private m_udtFunctions()    As TFunction
Private m_lngFncCnt         As Long

Private FunctionSinIndex    As Long
Private FunctionCosIndex    As Long
Private FunctionTanIndex    As Long
Private FunctionModIndex    As Long
Private FunctionLogIndex    As Long
Private FunctionSqrIndex    As Long
Private FunctionSgnIndex    As Long
Private FunctionAbsIndex    As Long

Private m_udtTokens()       As Token
Private m_lngTokensSize     As Long
Private m_lngPos            As Long

Private m_udtVars()         As Variable
Private m_lngVarCnt         As Long

Private m_strExpression     As String
Private m_udtExprArr()      As Integer
Private m_udtExprSA         As SAFEARRAY1D

Private m_intSValue(99)     As Integer

Private m_blnCalcError      As Boolean

Private m_blnError          As Boolean
Private m_strErrorMsg       As String

Public Property Let TokenList(udtTokens() As Token)
    m_udtTokens = udtTokens
    m_lngTokensSize = UBound(m_udtTokens) + 1
End Property

Public Property Get TokenList() As Token()
    TokenList = m_udtTokens
End Property

Public Property Get ExpressionString() As String
    ExpressionString = m_strExpression
End Property

Public Property Let ExpressionString(ByVal strExpr As String)
    Do While InStr(strExpr, " ") > 0: strExpr = Replace(strExpr, " ", ""): Loop
    m_strExpression = strExpr
    
    With m_udtExprSA
        .cElements = Len(m_strExpression) + 1
        .pvData = StrPtr(m_strExpression)
    End With
    
    CopySafeArray VarPtrArray(m_udtExprArr), VarPtr(m_udtExprSA), 4
    
    m_blnError = False
    m_strErrorMsg = vbNullString
    Tokenize
    
    DestroySafeArray VarPtrArray(m_udtExprArr), 4
End Property

Public Property Get ExpressionValid() As Boolean
    ExpressionValid = Not m_blnError
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = m_strErrorMsg
End Property

Public Function Evaluate() As Double
    If m_blnError Then
        Evaluate = 0
    Else
        m_lngPos = 1            ' skip TokenBegin
        m_blnCalcError = False
        Evaluate = Expression()
    End If
End Function

Public Property Get CalculationError() As Boolean
    CalculationError = m_blnCalcError
End Property

Private Function Expression() As Double
    Dim dblValue As Double

    dblValue = Term(OP_ANY_PREC)
    
    Do
        Select Case m_udtTokens(m_lngPos).ttype
            Case TokenAdd:
                m_lngPos = m_lngPos + 1
                dblValue = SafeAdd(dblValue, Term(OP_ADD_PREC))
            Case TokenSub:
                m_lngPos = m_lngPos + 1
                dblValue = SafeSub(dblValue, Term(OP_SUB_PREC))
            Case TokenEnd, TokenParR, TokenSep:
                Exit Do
        End Select
    Loop
    
    m_lngPos = m_lngPos + 1     ' skip "," and ")"
    
    Expression = dblValue
End Function

Private Function Term(ByVal OpPrec As Long) As Double
    Dim dblValue    As Double
    Dim lngFncIndex As Long
    
    ' Factor
    Select Case m_udtTokens(m_lngPos).ttype
        Case TokenFnc:
            lngFncIndex = m_udtTokens(m_lngPos).SyIdx
            m_lngPos = m_lngPos + 2     ' "Function", "("
            
            Select Case lngFncIndex
                Case FunctionCosIndex: dblValue = SafeCos(Expression())
                Case FunctionSinIndex: dblValue = SafeSin(Expression())
                Case FunctionTanIndex: dblValue = SafeTan(Expression())
                Case FunctionSqrIndex: dblValue = FncSqr(Expression())
                Case FunctionSgnIndex: dblValue = Sgn(Expression())
                Case FunctionAbsIndex: dblValue = Abs(Expression())
                Case FunctionLogIndex: dblValue = FncLog(Expression(), Expression())
                Case FunctionModIndex: dblValue = FncMod(Expression(), Expression())
            End Select
            
        Case TokenVar:
            dblValue = m_udtVars(m_udtTokens(m_lngPos).SyIdx).DValue
            m_lngPos = m_lngPos + 1
            
        Case TokenNumber:
            dblValue = m_udtTokens(m_lngPos).TDValue
            m_lngPos = m_lngPos + 1
            
        Case TokenParL:
            m_lngPos = m_lngPos + 1           ' "("
            dblValue = Expression()
            
    End Select
    
    Do While m_udtTokens(m_lngPos).OpPrec > OpPrec
        Select Case m_udtTokens(m_lngPos).ttype
            Case TokenPow:
                m_lngPos = m_lngPos + 1
                dblValue = SafePow(dblValue, Term(OP_POW_PREC))
            Case TokenMul:
                m_lngPos = m_lngPos + 1
                dblValue = SafeMul(dblValue, Term(OP_MUL_PREC))
            Case TokenDiv:
                m_lngPos = m_lngPos + 1
                dblValue = SafeDiv(dblValue, Term(OP_DIV_PREC))
            Case Else:
                Exit Do
        End Select
    Loop
    
    Term = dblValue
End Function

Private Function SafeTan(ByVal dblVal As Double) As Double
    On Error GoTo ErrorHandler
    
    SafeTan = Tan(dblVal)
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Private Function SafeCos(ByVal dblVal As Double) As Double
    On Error GoTo ErrorHandler
    
    SafeCos = Cos(dblVal)
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Private Function SafeSin(ByVal dblVal As Double) As Double
    On Error GoTo ErrorHandler
    
    SafeSin = Sin(dblVal)
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Private Function SafePow(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    On Error GoTo ErrorHandler
    
    SafePow = dblVal1 ^ dblVal2
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Private Function SafeDiv(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    On Error GoTo ErrorHandler
    
    SafeDiv = dblVal1 / dblVal2
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Private Function SafeMul(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    On Error GoTo ErrorHandler
    
    SafeMul = dblVal1 * dblVal2
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Private Function SafeSub(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    On Error GoTo ErrorHandler
    
    SafeSub = dblVal1 - dblVal2
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Private Function SafeAdd(ByVal dblVal1 As Double, ByVal dblVal2 As Double) As Double
    On Error GoTo ErrorHandler
    
    SafeAdd = dblVal1 + dblVal2
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Public Function ValidateExpression() As Boolean
    Dim lngBrcCnt   As Long
    
    ValidateExpression = ValidateLayer(1, 0, lngBrcCnt)
    
    If Not m_blnError Then
        If lngBrcCnt > 0 Then
            m_blnError = True
            m_strErrorMsg = "Expected: )"
            ValidateExpression = False
        End If
    End If
End Function

Private Function ValidateLayer(ByRef i As Long, ByVal lngParams As Long, ByRef lngBrcCnt As Long) As Boolean
    Dim udeExpected     As Tokens
    Dim lngParamCnt     As Long
    Dim lngNextParamCnt As Long
    
    udeExpected = TokenSub Or TokenAdd Or TokenEnd Or TokenFnc Or TokenNumber Or TokenParL Or TokenVar
    
    Do
        If (m_udtTokens(i).ttype And udeExpected) = 0 Then
            m_blnError = True
            m_strErrorMsg = "Unexpected: " & TokenToString(m_udtTokens(i).ttype)
            Exit Function
        End If
        
        Select Case m_udtTokens(i).ttype
            Case TokenBegin:
                udeExpected = TokenSub Or TokenAdd Or TokenEnd Or TokenFnc Or TokenNumber Or TokenParL Or TokenVar
                
            Case TokenAdd, TokenSub, TokenMul, TokenDiv, TokenPow:
                udeExpected = TokenFnc Or TokenNumber Or TokenParL Or TokenVar
                
            Case TokenFnc:
                udeExpected = TokenParL
                If m_udtTokens(i).SyIdx = -1 Then
                    m_blnError = True
                    m_strErrorMsg = "Unknown function: " & m_udtTokens(i).TValue
                    Exit Function
                Else
                    lngNextParamCnt = m_udtFunctions(m_udtTokens(i).SyIdx).params - 1
                End If
                
            Case TokenNumber:
                udeExpected = TOKEN_OP_FLAG Or TokenEnd Or TokenParR Or TokenSep
                
            Case TokenVar:
                udeExpected = TOKEN_OP_FLAG Or TokenEnd Or TokenParR Or TokenSep
                If m_udtTokens(i).SyIdx = -1 Then
                    m_blnError = True
                    m_strErrorMsg = "Unknown variable: " & m_udtTokens(i).TValue
                    Exit Function
                End If
                
            Case TokenParL:
                'udeExpected = TokenAdd Or TokenSub Or TokenFnc Or TokenNumber Or TokenParL Or TokenVar Or TokenParR
                udeExpected = TOKEN_OP_FLAG Or TokenParR Or TokenEnd Or TokenSep
                lngBrcCnt = lngBrcCnt + 1
                i = i + 1
                If Not ValidateLayer(i, lngNextParamCnt, lngBrcCnt) Then
                    Exit Function
                Else
                    lngNextParamCnt = 0
                End If
                
            Case TokenParR:
                udeExpected = TOKEN_OP_FLAG Or TokenParR Or TokenEnd Or TokenSep
                lngBrcCnt = lngBrcCnt - 1
                If lngBrcCnt = -1 Then
                    m_blnError = True
                    m_strErrorMsg = "too many closed brackets"
                End If
                Exit Do
                
            Case TokenSep:
                udeExpected = TokenAdd Or TokenSub Or TokenNumber Or TokenFnc Or TokenVar Or TokenParL
                lngParamCnt = lngParamCnt + 1
                
            Case TokenEnd:
                Exit Do
                
        End Select
        
        i = i + 1
    Loop While m_udtTokens(i).ttype <> TokenEnd
    
    If lngParamCnt <> lngParams Then
        m_blnError = True
        m_strErrorMsg = "wrong number of arguments"
    Else
        Select Case m_udtTokens(i - 1).ttype
            Case TokenAdd, TokenSub, TokenDiv, TokenMul, TokenPow, TokenFnc, TokenSep
                m_blnError = True
                m_strErrorMsg = "unexpected end"
        End Select
    End If
    
    If Not m_blnError Then
        ValidateLayer = True
    End If
End Function

Private Function TokenToString(ByVal tk As Tokens) As String
    Select Case tk
        Case TokenBegin:    TokenToString = "begin"
        Case TokenAdd:      TokenToString = "+"
        Case TokenDiv:      TokenToString = "/"
        Case TokenEnd:      TokenToString = "end"
        Case TokenFnc:      TokenToString = "function"
        Case TokenMul:      TokenToString = "*"
        Case TokenNumber:   TokenToString = "value"
        Case TokenParL:     TokenToString = "("
        Case TokenParR:     TokenToString = ")"
        Case TokenPow:      TokenToString = "^"
        Case TokenSep:      TokenToString = ","
        Case TokenSub:      TokenToString = "-"
        Case TokenVar:      TokenToString = "variable"
        Case Else:          TokenToString = "unknown token"
    End Select
End Function

Private Sub Tokenize()
    Dim lngTokenCnt     As Long
    Dim lngTermLen      As Long
    Dim i               As Long
    Dim j               As Long
    Dim lngSLen         As Long
    Dim pFirstSValElem  As Long

    pFirstSValElem = VarPtr(m_intSValue(0))

    m_udtTokens(lngTokenCnt).ttype = TokenBegin
    lngTokenCnt = lngTokenCnt + 1
    
    lngTermLen = UBound(m_udtExprArr)

    For i = 0 To lngTermLen - 1
        Select Case m_udtExprArr(i)
            Case CHAR_MINUS:
                Select Case m_udtExprArr(i + 1)
                    Case CHAR_PLUS:     ' -+
                        m_udtExprArr(i + 1) = CHAR_WS
                    Case CHAR_MINUS:    ' --
                        m_udtExprArr(i + 1) = CHAR_WS
                        m_udtExprArr(i) = CHAR_PLUS
                End Select
            Case CHAR_PLUS:
                Select Case m_udtExprArr(i + 1)
                    Case CHAR_PLUS:     ' ++
                        m_udtExprArr(i + 1) = CHAR_WS
                    Case CHAR_MINUS:    ' +-
                        m_udtExprArr(i) = CHAR_WS
                End Select
        End Select
    Next
    
    For i = 0 To lngTermLen
        Select Case m_udtExprArr(i)
        
            Case CHAR_PLUS:
                With m_udtTokens(lngTokenCnt)
                    .ttype = TokenAdd
                    .OpPrec = OP_ADD_PREC
                End With
                lngTokenCnt = lngTokenCnt + 1
                
            Case CHAR_MINUS:
                With m_udtTokens(lngTokenCnt)
                    .ttype = TokenSub
                    .OpPrec = OP_SUB_PREC
                End With
                lngTokenCnt = lngTokenCnt + 1
                
            Case CHAR_ASTERISK:
                With m_udtTokens(lngTokenCnt)
                    .ttype = TokenMul
                    .OpPrec = OP_MUL_PREC
                End With
                lngTokenCnt = lngTokenCnt + 1
                
            Case CHAR_SLASH:
                With m_udtTokens(lngTokenCnt)
                    .ttype = TokenDiv
                    .OpPrec = OP_DIV_PREC
                End With
                lngTokenCnt = lngTokenCnt + 1
                
            Case CHAR_POWER:
                With m_udtTokens(lngTokenCnt)
                    .ttype = TokenPow
                    .OpPrec = OP_POW_PREC
                End With
                lngTokenCnt = lngTokenCnt + 1
                
            Case CHAR_SEP:
                m_udtTokens(lngTokenCnt).ttype = TokenSep
                lngTokenCnt = lngTokenCnt + 1
                
            Case CHAR_PARL:
                m_udtTokens(lngTokenCnt).ttype = TokenParL
                lngTokenCnt = lngTokenCnt + 1
                
            Case CHAR_PARR:
                m_udtTokens(lngTokenCnt).ttype = TokenParR
                lngTokenCnt = lngTokenCnt + 1
                
            Case CHAR_WS, 0:
                ' ignore
            
            Case 48 To 57:
                m_intSValue(0) = m_udtExprArr(i)
                lngSLen = 1
                i = i + 1
                
                For j = i To lngTermLen
                    Select Case m_udtExprArr(j)
                        Case 48 To 57, 46:
                            m_intSValue(lngSLen) = m_udtExprArr(j)
                            lngSLen = lngSLen + 1
                        Case Else:
                            Exit For
                    End Select
                Next
                
                i = j - 1

                With m_udtTokens(lngTokenCnt)
                    .ttype = TokenNumber
                    .TDValue = Val(SysAllocStringLenPtr(pFirstSValElem, lngSLen))
                End With
                lngTokenCnt = lngTokenCnt + 1
                    
            Case 65 To 90, 97 To 122:
                m_intSValue(0) = m_udtExprArr(i)
                lngSLen = 1
                i = i + 1

                For j = i To lngTermLen
                    Select Case m_udtExprArr(j)
                        Case 48 To 57, 65 To 90, 97 To 122, 95:
                            m_intSValue(lngSLen) = m_udtExprArr(j)
                            lngSLen = lngSLen + 1
                        Case Else:
                            Exit For
                    End Select
                Next
                
                i = j - 1

                With m_udtTokens(lngTokenCnt)
                    .TValue = SysAllocStringLenPtr(pFirstSValElem, lngSLen)
                    If m_udtExprArr(i + 1) = CHAR_PARL Then
                        .ttype = TokenFnc
                        .SyIdx = GetFncIndex(.TValue)
                    Else
                        .ttype = TokenVar
                        .SyIdx = GetVariableIndex(.TValue)
                    End If
                End With
                lngTokenCnt = lngTokenCnt + 1

            Case Else:
                m_udtTokens(lngTokenCnt).ttype = TokenUnknown
                lngTokenCnt = lngTokenCnt + 1
                
        End Select
        
        If lngTokenCnt >= m_lngTokensSize Then
            m_lngTokensSize = m_lngTokensSize * 2
            ReDim Preserve m_udtTokens(m_lngTokensSize - 1) As Token
        End If
    Next
    
    m_udtTokens(lngTokenCnt).ttype = TokenEnd
End Sub

Private Function FncSqr(ByVal lhs As Double) As Double
    If lhs >= 0 Then
        On Error GoTo ErrorHandler
        FncSqr = Sqr(lhs)
        On Error GoTo 0
    Else
        m_blnCalcError = True
    End If
    
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Private Function FncLog(ByVal lhs As Double, ByVal rhs As Double) As Double
    If lhs <= 0 Or rhs <= 0 Then
        m_blnCalcError = True
    Else
        On Error GoTo ErrorHandler
        FncLog = Log(lhs) / Log(rhs)
        On Error GoTo 0
    End If
    
    Exit Function

ErrorHandler:
    m_blnCalcError = True
End Function

Private Function FncMod(ByVal lhs As Double, ByVal rhs As Double) As Double
    If Round(rhs, 0) <> 0 Then
        On Error GoTo ErrorHandler
        FncMod = lhs Mod rhs
        On Error GoTo 0
    Else
        m_blnCalcError = True
    End If
    
    Exit Function
    
ErrorHandler:
    m_blnCalcError = True
End Function

Public Property Get VariableName(ByVal index As Long) As String
    VariableName = m_udtVars(index).VarName
End Property

Public Property Get VariableValue(ByVal index As Long) As Double
    VariableValue = m_udtVars(index).DValue
End Property

Public Property Let VariableValue(ByVal index As Long, ByVal dblValue As Double)
    m_udtVars(index).DValue = dblValue
End Property

Public Function GetVariableIndex(ByVal name As String) As Long
    Dim i   As Long
    
    For i = 0 To m_lngVarCnt - 1
        If StrComp(m_udtVars(i).VarName, name, vbTextCompare) = 0 Then
            GetVariableIndex = i
            Exit Function
        End If
    Next
    
    GetVariableIndex = -1
End Function

Public Sub AddVariable(ByVal value As Double, ByVal name As String)
    ReDim Preserve m_udtVars(m_lngVarCnt) As Variable
    
    With m_udtVars(m_lngVarCnt)
        .DValue = value
        .VarName = name
    End With
    
    m_lngVarCnt = m_lngVarCnt + 1
End Sub

Private Function GetFncIndex(ByVal name As String) As Long
    Dim i   As Long
    
    For i = 0 To m_lngFncCnt - 1
        If StrComp(m_udtFunctions(i).name, name, vbTextCompare) = 0 Then
            GetFncIndex = i
            Exit Function
        End If
    Next
    
    GetFncIndex = -1
End Function

Private Function AddFnc(ByVal name As String, ByVal params As Long) As Long
    ReDim Preserve m_udtFunctions(m_lngFncCnt) As TFunction
    
    With m_udtFunctions(m_lngFncCnt)
        .params = params
        .name = name
    End With
    
    AddFnc = m_lngFncCnt
    m_lngFncCnt = m_lngFncCnt + 1
End Function

Public Sub Formula_Init()
    FunctionSinIndex = AddFnc("sin", 1)
    FunctionCosIndex = AddFnc("cos", 1)
    FunctionTanIndex = AddFnc("tan", 1)
    FunctionSqrIndex = AddFnc("sqr", 1)
    FunctionSgnIndex = AddFnc("sgn", 1)
    FunctionAbsIndex = AddFnc("abs", 1)
    FunctionModIndex = AddFnc("mod", 2)
    FunctionLogIndex = AddFnc("log", 2)
    
    AddVariable Atn(1) * 4, "pi"
    AddVariable Exp(1), "e"

    m_lngTokensSize = 50
    ReDim m_udtTokens(m_lngTokensSize - 1) As Token
    
    m_udtTokens(0).ttype = TokenBegin
    m_udtTokens(1).ttype = TokenEnd
    
    With m_udtExprSA
        .cbElements = 2
        .cDims = 1
        .pvData = StrPtr(m_strExpression)
        .cElements = Len(m_strExpression)
    End With
End Sub

Public Sub Formula_Terminate()
    '
End Sub
