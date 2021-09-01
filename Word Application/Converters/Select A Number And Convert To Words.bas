Attribute VB_Name = "NumToWords"
Option Explicit
                                                                        
'=====================================================================\
'Module Name: NumberToWords Translator                                |\
'Version: 1.0                                                         ||\
'Language: Hindi, Gujarati, English                                   |||\
'Designed & Developed By: KAMAL C. BHARAKHDA                          |||/
'Support & Testing By: O.P. Agrawal                                   ||/
'gitHub: https://github.com/vbamagician                               |/
'=====================================================================/

'Function Level Settings
Const strSnglSpc As String = " "
Const strDashChar As String = "-"
Const strInvalidMsg As String = "Error: Invalid Input"
Const strNumberLimitMsg As String = "Please select a Number between 0 to 99,99,99,99,999.99"
Const DefaultVal As Byte = 0
Const DigitsAfterDecimal As Byte = 2

'Language Constants
Const HindiChar As Byte = 0
Const EnglishChar As Byte = 1
Const GujaratiChar As Byte = 2

'Currency Form
Const CurrencyForm As Byte = 0
Const NumberForm As Byte = 1

Const strArab As String = "Arab"
Const strCrore As String = "Crore"
Const strLac As String = "Lakh"
Const strThousand As String = "Thousand"
Const strHundred As String = "Hundred"
Const strRupees As String = "Rupees"
Const strPaise As String = "Paise"
Const strAnd As String = "And"
Const strOnly As String = "Only"
Const strDecimal As String = "Decimal"
Const strRupee As String = "Rupee"
Const strPaisa As String = "Paisa"

'Property Declaration for reducing complexity in coding
'Input Currency Mapper
Private mInputCurrency As Byte
'NumberFormat or Not?
Private mIsCurrencyString As Byte

Private Property Get pInputCurrency() As Byte
    pInputCurrency = mInputCurrency
End Property

Private Property Let pInputCurrency(ByVal vNewValue As Byte)
    mInputCurrency = vNewValue
End Property

Private Property Get pIsCurrencyString() As Byte
    pIsCurrencyString = mIsCurrencyString
End Property

Private Property Let pIsCurrencyString(ByVal vNewValue As Byte)
    mIsCurrencyString = vNewValue
End Property

Public Sub NumberToWords()
    
    Dim InputLanguage As Long
    Dim MyNumber As Double
    Dim tmp As Variant
    Dim Sel As Selection
    Dim InputString As String
    Dim OutputString As String
    
    'Check if the selected text is number or not? not a number then it will exit the sub
    Set Sel = Application.Selection
    
    If Sel.Type <> wdSelectionIP Then
        InputString = VBA.Trim(Sel.Text)
        If Not VBA.IsNumeric(InputString) Then
            MsgBox "Selected text: '" & InputString & "' is not a Number.", vbCritical, "Selection Error"
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    'initializing the operating variables
    InputLanguage = 1
    tmp = VBA.InputBox("For, English insert 1," & vbNewLine & "For, Hindi insert 2," & vbNewLine & "For, Gujarati insert 3", "Choose Conversion Language", InputLanguage)
    
    'checking type of the input. if not numerical input then exit sub else, default will be english as 1
    If VBA.IsNumeric(tmp) Then
        If tmp >= 1 And tmp <= 3 Then
            InputLanguage = tmp
        End If
    Else
        Exit Sub
    End If
    
    'getting selected number and formatting it in number and storing it in double var.
    MyNumber = VBA.CDbl(InputString)
    
    'iscurrency?
    pIsCurrencyString = NumberForm
    
    Select Case InputLanguage
        Case 1
            'English
            pInputCurrency = EnglishChar
            OutputString = Translator(MyNumber)
            
        Case 2
            'Hindi
            pInputCurrency = HindiChar
            OutputString = Translator(MyNumber)
        
        Case 3
            'Gujarati
            pInputCurrency = GujaratiChar
            OutputString = Translator(MyNumber)
        
    End Select
    
    'Result
    If OutputString <> vbNullString Then Sel.Range.Text = VBA.Trim(OutputString)
     
End Sub

Private Function Translator(ByVal MyNumber As Variant, _
                            Optional opt_CurrencyStringPosition As Variant = DefaultVal) As String
    
    'Declaration of Supporting Members
    Dim DecimalPlace As Long, ActNum As Double
    Dim myCrore As Long, myLac As Long, myThousand As Long
    Dim MyHundred As Long, myTen As Long, myPaise As Long
    Dim myAr As String, myCr As String, myLc As String, myTh As String, myArab As String
    Dim MyHd As String, myTn As String, myPai As String
    Dim MyRu As String, MyAnd As String, MyOnly As String, MyPais As String
    Dim FinalString As String, FlagStatus As Byte
    
    'Input validation
    If VBA.IsNumeric(MyNumber) = False Or MyNumber = vbNullString Then
        Translator = strInvalidMsg
        Exit Function
    Else
        If MyNumber < 0 Then
            MyNumber = VBA.Abs(MyNumber)
        End If
        If MyNumber > 99999999999.99 Then
            Translator = strNumberLimitMsg
            Exit Function
        End If
        If MyNumber = 0 Then
            Translator = LoopkupHString(MyNumber)
            Exit Function
        End If
    End If
    
    'Validation for Optional Arguments
    If VBA.IsNumeric(opt_CurrencyStringPosition) = False Then
        opt_CurrencyStringPosition = DefaultVal
    End If
    
    'Storing Actual Value to Another variable for final logics.
    ActNum = VBA.Round(VBA.CDbl(MyNumber), DigitsAfterDecimal)
    
    ' String representation of amount.
    If MyNumber > 1 And MyNumber < 0 Then
        MyNumber = Trim(Str(MyNumber))
    End If
    
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(MyNumber, Application.International(wdDecimalSeparator))
    
    ' Convert Paise and set MyNumber to Rupees amount.
    If DecimalPlace > 0 Then
        myPaise = Left(Mid(VBA.Round(MyNumber, DigitsAfterDecimal), DecimalPlace + 1) & "00", DigitsAfterDecimal)
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    
    'Sorting and Saggrigation
    myArab = RoundDownNumber(MyNumber / 1000000000)
    myCrore = RoundDownNumber((MyNumber - (myArab * 1000000000)) / 10000000)
    myLac = RoundDownNumber((MyNumber - (myArab * 1000000000) - (myCrore * 10000000)) / 100000)
    myThousand = RoundDownNumber((MyNumber - (myArab * 1000000000) - (myCrore * 10000000) - (myLac * 100000)) / 1000)
    MyHundred = RoundDownNumber((MyNumber - (myArab * 1000000000) - (myCrore * 10000000) - (myLac * 100000) - (myThousand * 1000)) / 100)
    myTen = RoundDownNumber((MyNumber - (myArab * 1000000000) - (myCrore * 10000000) - (myLac * 100000) - (myThousand * 1000) - (MyHundred * 100)))
    
    'Construction
    myAr = IIf(myArab > 0, LoopkupHString(myArab) & strSnglSpc & LoopkupHString(strArab) & strSnglSpc, vbNullString)
    myCr = IIf(myCrore > 0, LoopkupHString(myCrore) & strSnglSpc & LoopkupHString(strCrore) & strSnglSpc, vbNullString)
    myLc = IIf(myLac > 0, LoopkupHString(myLac) & strSnglSpc & LoopkupHString(strLac) & strSnglSpc, vbNullString)
    myTh = IIf(myThousand > 0, LoopkupHString(myThousand) & strSnglSpc & LoopkupHString(strThousand) & strSnglSpc, vbNullString)
    MyHd = IIf(MyHundred > 0, LoopkupHString(MyHundred) & strSnglSpc & LoopkupHString(strHundred) & strSnglSpc, vbNullString)
    myTn = IIf(myTen > 0, LoopkupHString(myTen) & strSnglSpc, vbNullString)
    myPai = IIf(DecimalPlace > 0, LoopkupHString(myPaise) & strSnglSpc, vbNullString)
    
    'Currency Parameters Construction
    If pIsCurrencyString = CurrencyForm Then
        If ActNum = 1 Then
            MyRu = LoopkupHString(strRupee) & strSnglSpc
            MyOnly = LoopkupHString(strOnly) & strSnglSpc
        ElseIf ActNum > 0 And ActNum < 1 Then
            MyRu = vbNullString
        ElseIf ActNum > 1 And ActNum < 2 Then
            MyRu = LoopkupHString(strRupee) & strSnglSpc
            If DecimalPlace > 0 Then
                MyAnd = LoopkupHString(strAnd) & strSnglSpc
            Else
                PlaceAND myAr, myCr, myLc, myTh, MyHd, myTn
            End If
        Else
            MyRu = LoopkupHString(strRupees) & strSnglSpc
            If DecimalPlace > 0 Then
                MyAnd = LoopkupHString(strAnd) & strSnglSpc
            Else
                PlaceAND myAr, myCr, myLc, myTh, MyHd, myTn
            End If
        End If
        MyPais = IIf(DecimalPlace > 0, IIf(myPaise = 1, LoopkupHString(strPaisa) & strSnglSpc, LoopkupHString(strPaise) & strSnglSpc), vbNullString)
        MyOnly = IIf(DecimalPlace > 0, vbNullString, LoopkupHString(strOnly) & strSnglSpc)
    Else
        MyRu = IIf(ActNum > 0 And ActNum < 1, LoopkupHString(0) & strSnglSpc, vbNullString)
        If DecimalPlace > 0 Then
            If VBA.Len(myPaise) = 2 Then
                myPai = LoopkupHString(VBA.Left(myPaise, 1)) & IIf(VBA.Right(myPaise, 1) = 0, vbNullString, strSnglSpc & LoopkupHString(VBA.Right(myPaise, 1)))
            ElseIf VBA.Len(myPaise) = 1 Then
                myPai = LoopkupHString(0) & strSnglSpc & LoopkupHString(myPaise)
            End If
        Else
            myPai = vbNullString
        End If
        If DecimalPlace > 0 Then
            MyAnd = LoopkupHString(strDecimal) & strSnglSpc
        Else
            PlaceAND myAr, myCr, myLc, myTh, MyHd, myTn
        End If
    End If
    
    'Assembling
    If opt_CurrencyStringPosition = 0 Then
        Translator = myAr & myCr & myLc & myTh & MyHd & myTn & MyRu & MyAnd & myPai & MyPais & MyOnly
    ElseIf opt_CurrencyStringPosition = 1 Then
        Translator = MyRu & myAr & myCr & myLc & myTh & MyHd & myTn & MyAnd & MyPais & myPai & MyOnly
    Else
        Translator = myAr & myCr & myLc & myTh & MyHd & myTn & MyRu & MyAnd & myPai & MyPais & MyOnly
    End If
    
    'Exit condition if something doesn't work well
    If Translator = vbNullString Then
        Translator = strDashChar
    End If
    
End Function

Private Function RoundDownNumber(ByVal MyNumber As String) As Integer
    Dim DecimalPlace As Byte
    DecimalPlace = VBA.InStr(1, MyNumber, Application.International(wdDecimalSeparator))
    If DecimalPlace > 0 Then
        RoundDownNumber = VBA.Mid(MyNumber, 1, DecimalPlace - 1)
    Else
        RoundDownNumber = MyNumber
    End If
End Function

Private Sub PlaceAND(ByRef myAr As String, _
                     ByRef myCr As String, _
                     ByRef myLc As String, _
                     ByRef myTh As String, _
                     ByRef MyHd As String, _
                     ByRef myTn As String)
    
    'Should it be work or not!
    If pInputCurrency <> EnglishChar Then Exit Sub
    
    'myAr and myTn will not be change through out whether it will have value or not
    If myAr <> vbNullString Then
        If myCr <> vbNullString And myLc = vbNullString And myTh = vbNullString And MyHd = vbNullString And myTn = vbNullString Then
            myCr = LoopkupHString(strAnd) & strSnglSpc & myCr
        ElseIf myCr = vbNullString And myLc <> vbNullString And myTh = vbNullString And MyHd = vbNullString And myTn = vbNullString Then
            myLc = LoopkupHString(strAnd) & strSnglSpc & myLc
        ElseIf myCr = vbNullString And myLc = vbNullString And myTh <> vbNullString And MyHd = vbNullString And myTn = vbNullString Then
            myTh = LoopkupHString(strAnd) & strSnglSpc & myTh
        ElseIf myCr = vbNullString And myLc = vbNullString And myTh = vbNullString And MyHd <> vbNullString And myTn = vbNullString Then
            MyHd = LoopkupHString(strAnd) & strSnglSpc & MyHd
        ElseIf myCr = vbNullString And myLc = vbNullString And myTh = vbNullString And MyHd = vbNullString And myTn <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        ElseIf MyHd <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        End If
    ElseIf myCr <> vbNullString Then
        If myLc <> vbNullString And myTh = vbNullString And MyHd = vbNullString And myTn = vbNullString Then
            myLc = LoopkupHString(strAnd) & strSnglSpc & myLc
        ElseIf myLc = vbNullString And myTh <> vbNullString And MyHd = vbNullString And myTn = vbNullString Then
            myTh = LoopkupHString(strAnd) & strSnglSpc & myTh
        ElseIf myLc = vbNullString And myTh = vbNullString And MyHd <> vbNullString And myTn = vbNullString Then
            MyHd = LoopkupHString(strAnd) & strSnglSpc & MyHd
        ElseIf myLc = vbNullString And myTh = vbNullString And MyHd = vbNullString And myTn <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        ElseIf MyHd <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        End If
    ElseIf myLc <> vbNullString Then
        If myTh <> vbNullString And MyHd = vbNullString And myTn = vbNullString Then
            myTh = LoopkupHString(strAnd) & strSnglSpc & myTh
        ElseIf myTh = vbNullString And MyHd <> vbNullString And myTn = vbNullString Then
            MyHd = LoopkupHString(strAnd) & strSnglSpc & MyHd
        ElseIf myTh = vbNullString And MyHd = vbNullString And myTn <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        ElseIf MyHd <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        End If
    ElseIf myTh <> vbNullString Then
        If MyHd <> vbNullString And myTn = vbNullString Then
            MyHd = LoopkupHString(strAnd) & strSnglSpc & MyHd
        ElseIf MyHd = vbNullString And myTn <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        ElseIf myTn <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        End If
    ElseIf MyHd <> vbNullString Then
        If myTn <> vbNullString Then
            myTn = LoopkupHString(strAnd) & strSnglSpc & myTn
        End If
    End If
    
End Sub

Private Function LoopkupHString(ByVal MyValue As Variant) As String

    If pInputCurrency = HindiChar Then 'Hindi Translation in Hindi Characters
        Select Case MyValue
            Case "Hundred": LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2380)), vbNullString)
            Case "Thousand": LoopkupHString = VBA.Join(Array(VBA.ChrW(2361), VBA.ChrW(2332), VBA.ChrW(2364), VBA.ChrW(2366), VBA.ChrW(2352)), vbNullString)
            Case "Lakh": LoopkupHString = VBA.Join(Array(VBA.ChrW(2354), VBA.ChrW(2366), VBA.ChrW(2326)), vbNullString)
            Case "Crore": LoopkupHString = VBA.Join(Array(VBA.ChrW(2325), VBA.ChrW(2352), VBA.ChrW(2379), VBA.ChrW(2337), VBA.ChrW(2364)), vbNullString)
            Case "Arab": LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2352), VBA.ChrW(2348)), vbNullString)
            Case "Kharab": LoopkupHString = VBA.Join(Array(VBA.ChrW(2326), VBA.ChrW(2352), VBA.ChrW(2348)), vbNullString)
            Case "Rupees": LoopkupHString = VBA.Join(Array(VBA.ChrW(2352), VBA.ChrW(2369), VBA.ChrW(2346), VBA.ChrW(2351), VBA.ChrW(2375)), vbNullString)
            Case "Paise": LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2376), VBA.ChrW(2360), VBA.ChrW(2375)), vbNullString)
            Case "Rupee": LoopkupHString = VBA.Join(Array(VBA.ChrW(2352), VBA.ChrW(2369), VBA.ChrW(2346), VBA.ChrW(2351), VBA.ChrW(2366)), vbNullString)
            Case "Paisa": LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2376), VBA.ChrW(2360), VBA.ChrW(2366)), vbNullString)
            Case "And": LoopkupHString = VBA.Join(Array(VBA.ChrW(2324), VBA.ChrW(2352)), vbNullString)
            Case "Only": LoopkupHString = VBA.Join(Array(VBA.ChrW(2350), VBA.ChrW(2366), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2352)), vbNullString)
            Case "Decimal": LoopkupHString = VBA.Join(Array(VBA.ChrW(2342), VBA.ChrW(2358), VBA.ChrW(2350), VBA.ChrW(2354), VBA.ChrW(2357)), vbNullString)
            Case 0: LoopkupHString = VBA.Join(Array(VBA.ChrW(2358), VBA.ChrW(2370), VBA.ChrW(2344), VBA.ChrW(2381), VBA.ChrW(2351)), vbNullString)
            Case 1: LoopkupHString = VBA.Join(Array(VBA.ChrW(2319), VBA.ChrW(2325)), vbNullString)
            Case 2: LoopkupHString = VBA.Join(Array(VBA.ChrW(2342), VBA.ChrW(2379)), vbNullString)
            Case 3: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2344)), vbNullString)
            Case 4: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2366), VBA.ChrW(2352)), vbNullString)
            Case 5: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2366), VBA.ChrW(2306), VBA.ChrW(2330)), vbNullString)
            Case 6: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2361)), vbNullString)
            Case 7: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2366), VBA.ChrW(2340)), vbNullString)
            Case 8: LoopkupHString = VBA.Join(Array(VBA.ChrW(2310), VBA.ChrW(2336)), vbNullString)
            Case 9: LoopkupHString = VBA.Join(Array(VBA.ChrW(2344), VBA.ChrW(2380)), vbNullString)
            Case 10: LoopkupHString = VBA.Join(Array(VBA.ChrW(2342), VBA.ChrW(2360)), vbNullString)
            Case 11: LoopkupHString = VBA.Join(Array(VBA.ChrW(2327), VBA.ChrW(2381), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2352), VBA.ChrW(2361)), vbNullString)
            Case 12: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2366), VBA.ChrW(2352), VBA.ChrW(2361)), vbNullString)
            Case 13: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2375), VBA.ChrW(2352), VBA.ChrW(2361)), vbNullString)
            Case 14: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2380), VBA.ChrW(2342), VBA.ChrW(2361)), vbNullString)
            Case 15: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2306), VBA.ChrW(2342), VBA.ChrW(2381), VBA.ChrW(2352), VBA.ChrW(2361)), vbNullString)
            Case 16: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2379), VBA.ChrW(2354), VBA.ChrW(2361)), vbNullString)
            Case 17: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2352), VBA.ChrW(2361)), vbNullString)
            Case 18: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2336), VBA.ChrW(2366), VBA.ChrW(2352), VBA.ChrW(2361)), vbNullString)
            Case 19: LoopkupHString = VBA.Join(Array(VBA.ChrW(2313), VBA.ChrW(2344), VBA.ChrW(2381), VBA.ChrW(2344), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 20: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 21: LoopkupHString = VBA.Join(Array(VBA.ChrW(2311), VBA.ChrW(2325), VBA.ChrW(2381), VBA.ChrW(2325), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 22: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2366), VBA.ChrW(2312), VBA.ChrW(2360)), vbNullString)
            Case 23: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2375), VBA.ChrW(2312), VBA.ChrW(2360)), vbNullString)
            Case 24: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2380), VBA.ChrW(2348), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 25: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2330), VBA.ChrW(2381), VBA.ChrW(2330), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 26: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2348), VBA.ChrW(2381), VBA.ChrW(2348), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 27: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2312), VBA.ChrW(2360)), vbNullString)
            Case 28: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2335), VBA.ChrW(2381), VBA.ChrW(2336), VBA.ChrW(2366), VBA.ChrW(2311), VBA.ChrW(2360)), vbNullString)
            Case 29: LoopkupHString = VBA.Join(Array(VBA.ChrW(2313), VBA.ChrW(2344), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 30: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 31: LoopkupHString = VBA.Join(Array(VBA.ChrW(2311), VBA.ChrW(2325), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 32: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 33: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2375), VBA.ChrW(2306), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 34: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2380), VBA.ChrW(2306), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 35: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2376), VBA.ChrW(2306), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 36: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 37: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2376), VBA.ChrW(2306), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 38: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2337), VBA.ChrW(2364), VBA.ChrW(2340), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 39: LoopkupHString = VBA.Join(Array(VBA.ChrW(2313), VBA.ChrW(2344), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 40: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 41: LoopkupHString = VBA.Join(Array(VBA.ChrW(2311), VBA.ChrW(2325), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 42: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 43: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2376), VBA.ChrW(2306), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 44: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2357), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 45: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2376), VBA.ChrW(2306), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 46: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2367), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 47: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2376), VBA.ChrW(2306), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 48: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2337), VBA.ChrW(2364), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2354), VBA.ChrW(2368), VBA.ChrW(2360)), vbNullString)
            Case 49: LoopkupHString = VBA.Join(Array(VBA.ChrW(2313), VBA.ChrW(2344), VBA.ChrW(2330), VBA.ChrW(2366), VBA.ChrW(2360)), vbNullString)
            Case 50: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2330), VBA.ChrW(2366), VBA.ChrW(2360)), vbNullString)
            Case 51: LoopkupHString = VBA.Join(Array(VBA.ChrW(2311), VBA.ChrW(2325), VBA.ChrW(2381), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2357), VBA.ChrW(2344)), vbNullString)
            Case 52: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2366), VBA.ChrW(2357), VBA.ChrW(2344)), vbNullString)
            Case 53: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2367), VBA.ChrW(2352), VBA.ChrW(2346), VBA.ChrW(2344)), vbNullString)
            Case 54: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2380), VBA.ChrW(2357), VBA.ChrW(2344)), vbNullString)
            Case 55: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2330), VBA.ChrW(2346), VBA.ChrW(2344)), vbNullString)
            Case 56: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2346), VBA.ChrW(2381), VBA.ChrW(2346), VBA.ChrW(2344)), vbNullString)
            Case 57: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2357), VBA.ChrW(2344)), vbNullString)
            Case 58: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2335), VBA.ChrW(2381), VBA.ChrW(2336), VBA.ChrW(2366), VBA.ChrW(2357), VBA.ChrW(2344)), vbNullString)
            Case 59: LoopkupHString = VBA.Join(Array(VBA.ChrW(2313), VBA.ChrW(2344), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 60: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2366), VBA.ChrW(2336)), vbNullString)
            Case 61: LoopkupHString = VBA.Join(Array(VBA.ChrW(2311), VBA.ChrW(2325), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 62: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 63: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2367), VBA.ChrW(2352), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 64: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2380), VBA.ChrW(2306), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 65: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2376), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 66: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2367), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 67: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2337), VBA.ChrW(2364), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 68: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2337), VBA.ChrW(2364), VBA.ChrW(2360), VBA.ChrW(2336)), vbNullString)
            Case 69: LoopkupHString = VBA.Join(Array(VBA.ChrW(2313), VBA.ChrW(2344), VBA.ChrW(2361), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 70: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 71: LoopkupHString = VBA.Join(Array(VBA.ChrW(2311), VBA.ChrW(2325), VBA.ChrW(2361), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 72: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2361), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 73: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2367), VBA.ChrW(2361), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 74: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2380), VBA.ChrW(2361), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 75: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2330), VBA.ChrW(2361), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 76: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2367), VBA.ChrW(2361), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 77: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2340), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 78: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2336), VBA.ChrW(2361), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2352)), vbNullString)
            Case 79: LoopkupHString = VBA.Join(Array(VBA.ChrW(2313), VBA.ChrW(2344), VBA.ChrW(2381), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 80: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2360), VBA.ChrW(2381), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 81: LoopkupHString = VBA.Join(Array(VBA.ChrW(2311), VBA.ChrW(2325), VBA.ChrW(2381), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 82: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 83: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2367), VBA.ChrW(2352), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 84: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2380), VBA.ChrW(2352), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 85: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2330), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 86: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2367), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 87: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 88: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2336), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 89: LoopkupHString = VBA.Join(Array(VBA.ChrW(2344), VBA.ChrW(2357), VBA.ChrW(2366), VBA.ChrW(2360), VBA.ChrW(2368)), vbNullString)
            Case 90: LoopkupHString = VBA.Join(Array(VBA.ChrW(2344), VBA.ChrW(2357), VBA.ChrW(2381), VBA.ChrW(2357), VBA.ChrW(2375)), vbNullString)
            Case 91: LoopkupHString = VBA.Join(Array(VBA.ChrW(2311), VBA.ChrW(2325), VBA.ChrW(2381), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2348), VBA.ChrW(2375)), vbNullString)
            Case 92: LoopkupHString = VBA.Join(Array(VBA.ChrW(2348), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2381), VBA.ChrW(2357), VBA.ChrW(2375)), vbNullString)
            Case 93: LoopkupHString = VBA.Join(Array(VBA.ChrW(2340), VBA.ChrW(2367), VBA.ChrW(2352), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2357), VBA.ChrW(2375)), vbNullString)
            Case 94: LoopkupHString = VBA.Join(Array(VBA.ChrW(2330), VBA.ChrW(2380), VBA.ChrW(2352), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2357), VBA.ChrW(2375)), vbNullString)
            Case 95: LoopkupHString = VBA.Join(Array(VBA.ChrW(2346), VBA.ChrW(2306), VBA.ChrW(2330), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2357), VBA.ChrW(2375)), vbNullString)
            Case 96: LoopkupHString = VBA.Join(Array(VBA.ChrW(2331), VBA.ChrW(2367), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2348), VBA.ChrW(2375)), vbNullString)
            Case 97: LoopkupHString = VBA.Join(Array(VBA.ChrW(2360), VBA.ChrW(2340), VBA.ChrW(2381), VBA.ChrW(2340), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2357), VBA.ChrW(2375)), vbNullString)
            Case 98: LoopkupHString = VBA.Join(Array(VBA.ChrW(2309), VBA.ChrW(2336), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2357), VBA.ChrW(2375)), vbNullString)
            Case 99: LoopkupHString = VBA.Join(Array(VBA.ChrW(2344), VBA.ChrW(2367), VBA.ChrW(2344), VBA.ChrW(2381), VBA.ChrW(2351), VBA.ChrW(2366), VBA.ChrW(2344), VBA.ChrW(2357), VBA.ChrW(2375)), vbNullString)
        End Select
    ElseIf pInputCurrency = EnglishChar Then 'Indian Currency Format in English
        Select Case MyValue
            Case "Hundred": LoopkupHString = "Hundred"
            Case "Thousand": LoopkupHString = "Thousand"
            Case "Lakh": LoopkupHString = "Lakh"
            Case "Crore": LoopkupHString = "Crore"
            Case "Arab": LoopkupHString = "Arab"
            Case "Kharab": LoopkupHString = "Kharab"
            Case "Rupees": LoopkupHString = "Rupees"
            Case "Paise": LoopkupHString = "Paise"
            Case "Rupee": LoopkupHString = "Rupee"
            Case "Paisa": LoopkupHString = "Paisa"
            Case "And": LoopkupHString = "And"
            Case "Only": LoopkupHString = "Only"
            Case "Decimal": LoopkupHString = "Point"
            Case 0: LoopkupHString = "Zero"
            Case 1: LoopkupHString = "One"
            Case 2: LoopkupHString = "Two"
            Case 3: LoopkupHString = "Three"
            Case 4: LoopkupHString = "Four"
            Case 5: LoopkupHString = "Five"
            Case 6: LoopkupHString = "Six"
            Case 7: LoopkupHString = "Seven"
            Case 8: LoopkupHString = "Eight"
            Case 9: LoopkupHString = "Nine"
            Case 10: LoopkupHString = "Ten"
            Case 11: LoopkupHString = "Eleven"
            Case 12: LoopkupHString = "Twelve"
            Case 13: LoopkupHString = "Thirteen"
            Case 14: LoopkupHString = "Fourteen"
            Case 15: LoopkupHString = "Fifteen"
            Case 16: LoopkupHString = "Sixteen"
            Case 17: LoopkupHString = "Seventeen"
            Case 18: LoopkupHString = "Eighteen"
            Case 19: LoopkupHString = "Nineteen"
            Case 20: LoopkupHString = "Twenty"
            Case 21: LoopkupHString = "Twenty One"
            Case 22: LoopkupHString = "Twenty Two"
            Case 23: LoopkupHString = "Twenty Three"
            Case 24: LoopkupHString = "Twenty Four"
            Case 25: LoopkupHString = "Twenty Five"
            Case 26: LoopkupHString = "Twenty Six"
            Case 27: LoopkupHString = "Twenty Seven"
            Case 28: LoopkupHString = "Twenty Eight"
            Case 29: LoopkupHString = "Twenty Nine"
            Case 30: LoopkupHString = "Thirty"
            Case 31: LoopkupHString = "Thirty One"
            Case 32: LoopkupHString = "Thirty Two"
            Case 33: LoopkupHString = "Thirty Three"
            Case 34: LoopkupHString = "Thirty Four"
            Case 35: LoopkupHString = "Thirty Five"
            Case 36: LoopkupHString = "Thirty Six"
            Case 37: LoopkupHString = "Thirty Seven"
            Case 38: LoopkupHString = "Thirty Eight"
            Case 39: LoopkupHString = "Thirty Nine"
            Case 40: LoopkupHString = "Forty"
            Case 41: LoopkupHString = "Forty One"
            Case 42: LoopkupHString = "Forty Two"
            Case 43: LoopkupHString = "Forty Three"
            Case 44: LoopkupHString = "Forty Four"
            Case 45: LoopkupHString = "Forty Five"
            Case 46: LoopkupHString = "Forty Six"
            Case 47: LoopkupHString = "Forty Seven"
            Case 48: LoopkupHString = "Forty Eight"
            Case 49: LoopkupHString = "Forty Nine"
            Case 50: LoopkupHString = "Fifty"
            Case 51: LoopkupHString = "Fifty One"
            Case 52: LoopkupHString = "Fifty Two"
            Case 53: LoopkupHString = "Fifty Three"
            Case 54: LoopkupHString = "Fifty Four"
            Case 55: LoopkupHString = "Fifty Five"
            Case 56: LoopkupHString = "Fifty Six"
            Case 57: LoopkupHString = "Fifty Seven"
            Case 58: LoopkupHString = "Fifty Eight"
            Case 59: LoopkupHString = "Fifty Nine"
            Case 60: LoopkupHString = "Sixty"
            Case 61: LoopkupHString = "Sixty One"
            Case 62: LoopkupHString = "Sixty Two"
            Case 63: LoopkupHString = "Sixty Three"
            Case 64: LoopkupHString = "Sixty Four"
            Case 65: LoopkupHString = "Sixty Five"
            Case 66: LoopkupHString = "Sixty Six"
            Case 67: LoopkupHString = "Sixty Seven"
            Case 68: LoopkupHString = "Sixty Eight"
            Case 69: LoopkupHString = "Sixty Nine"
            Case 70: LoopkupHString = "Seventy"
            Case 71: LoopkupHString = "Seventy One"
            Case 72: LoopkupHString = "Seventy Two"
            Case 73: LoopkupHString = "Seventy Three"
            Case 74: LoopkupHString = "Seventy Four"
            Case 75: LoopkupHString = "Seventy Five"
            Case 76: LoopkupHString = "Seventy Six"
            Case 77: LoopkupHString = "Seventy Seven"
            Case 78: LoopkupHString = "Seventy Eight"
            Case 79: LoopkupHString = "Seventy Nine"
            Case 80: LoopkupHString = "Eighty"
            Case 81: LoopkupHString = "Eighty One"
            Case 82: LoopkupHString = "Eighty Two"
            Case 83: LoopkupHString = "Eighty Three"
            Case 84: LoopkupHString = "Eighty Four"
            Case 85: LoopkupHString = "Eighty Five"
            Case 86: LoopkupHString = "Eighty Six"
            Case 87: LoopkupHString = "Eighty Seven"
            Case 88: LoopkupHString = "Eighty Eight"
            Case 89: LoopkupHString = "Eighty Nine"
            Case 90: LoopkupHString = "Ninety"
            Case 91: LoopkupHString = "Ninety One"
            Case 92: LoopkupHString = "Ninety Two"
            Case 93: LoopkupHString = "Ninety Three"
            Case 94: LoopkupHString = "Ninety Four"
            Case 95: LoopkupHString = "Ninety Five"
            Case 96: LoopkupHString = "Ninety Six"
            Case 97: LoopkupHString = "Ninety Seven"
            Case 98: LoopkupHString = "Ninety Eight"
            Case 99: LoopkupHString = "Ninety Nine"
        End Select
    ElseIf pInputCurrency = GujaratiChar Then 'In gujarati
        Select Case MyValue
            Case "Hundred": LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2763)), vbNullString)
            Case "Thousand": LoopkupHString = VBA.Join(Array(VBA.ChrW(2745), VBA.ChrW(2716), VBA.ChrW(2750), VBA.ChrW(2736)), vbNullString)
            Case "Lakh": LoopkupHString = VBA.Join(Array(VBA.ChrW(2738), VBA.ChrW(2750), VBA.ChrW(2710)), vbNullString)
            Case "Crore": LoopkupHString = VBA.Join(Array(VBA.ChrW(2709), VBA.ChrW(2736), VBA.ChrW(2763), VBA.ChrW(2721)), vbNullString)
            Case "Arab": LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2736), VBA.ChrW(2732)), vbNullString)
            Case "Kharab": LoopkupHString = VBA.Join(Array(VBA.ChrW(2710), VBA.ChrW(2736), VBA.ChrW(2732)), vbNullString)
            Case "Rupees": LoopkupHString = VBA.Join(Array(VBA.ChrW(2736), VBA.ChrW(2754), VBA.ChrW(2730), VBA.ChrW(2751), VBA.ChrW(2735), VBA.ChrW(2750)), vbNullString)
            Case "Paise": LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2760), VBA.ChrW(2744), VBA.ChrW(2750)), vbNullString)
            Case "Rupee": LoopkupHString = VBA.Join(Array(VBA.ChrW(2736), VBA.ChrW(2754), VBA.ChrW(2730), VBA.ChrW(2751), VBA.ChrW(2735), VBA.ChrW(2750)), vbNullString)
            Case "Paisa": LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2760), VBA.ChrW(2744), VBA.ChrW(2763)), vbNullString)
            Case "And": LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2728), VBA.ChrW(2759)), vbNullString)
            Case "Only": LoopkupHString = VBA.Join(Array(VBA.ChrW(2734), VBA.ChrW(2750), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736)), vbNullString)
            Case "Decimal": LoopkupHString = VBA.Join(Array(VBA.ChrW(2726), VBA.ChrW(2742), VBA.ChrW(2750), VBA.ChrW(2690), VBA.ChrW(2742)), vbNullString)
            Case 0: LoopkupHString = VBA.Join(Array(VBA.ChrW(2742), VBA.ChrW(2754), VBA.ChrW(2728), VBA.ChrW(2765), VBA.ChrW(2735)), vbNullString)
            Case 1: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709)), vbNullString)
            Case 2: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2759)), vbNullString)
            Case 3: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2723)), vbNullString)
            Case 4: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2750), VBA.ChrW(2736)), vbNullString)
            Case 5: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2750), VBA.ChrW(2690), VBA.ChrW(2714)), vbNullString)
            Case 6: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715)), vbNullString)
            Case 7: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2750), VBA.ChrW(2724)), vbNullString)
            Case 8: LoopkupHString = VBA.Join(Array(VBA.ChrW(2694), VBA.ChrW(2720)), vbNullString)
            Case 9: LoopkupHString = VBA.Join(Array(VBA.ChrW(2728), VBA.ChrW(2741)), vbNullString)
            Case 10: LoopkupHString = VBA.Join(Array(VBA.ChrW(2726), VBA.ChrW(2744)), vbNullString)
            Case 11: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2711), VBA.ChrW(2751), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2736)), vbNullString)
            Case 12: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2750), VBA.ChrW(2736)), vbNullString)
            Case 13: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 14: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2764), VBA.ChrW(2726)), vbNullString)
            Case 15: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2690), VBA.ChrW(2726), VBA.ChrW(2736)), vbNullString)
            Case 16: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2763), VBA.ChrW(2739)), vbNullString)
            Case 17: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2724), VBA.ChrW(2736)), vbNullString)
            Case 18: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2722), VBA.ChrW(2750), VBA.ChrW(2736)), vbNullString)
            Case 19: LoopkupHString = VBA.Join(Array(VBA.ChrW(2707), VBA.ChrW(2711), VBA.ChrW(2723), VBA.ChrW(2751), VBA.ChrW(2744)), vbNullString)
            Case 20: LoopkupHString = VBA.Join(Array(VBA.ChrW(2741), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 21: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709), VBA.ChrW(2741), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 22: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2750), VBA.ChrW(2741), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 23: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2741), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 24: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2763), VBA.ChrW(2741), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 25: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2714), VBA.ChrW(2765), VBA.ChrW(2714), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 26: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715), VBA.ChrW(2741), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 27: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2741), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 28: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2720), VBA.ChrW(2765), VBA.ChrW(2720), VBA.ChrW(2750), VBA.ChrW(2741), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 29: LoopkupHString = VBA.Join(Array(VBA.ChrW(2707), VBA.ChrW(2711), VBA.ChrW(2723), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 30: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 31: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 32: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 33: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 34: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 35: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2750), VBA.ChrW(2690), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 36: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 37: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2721), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 38: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2721), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 39: LoopkupHString = VBA.Join(Array(VBA.ChrW(2707), VBA.ChrW(2711), VBA.ChrW(2723), VBA.ChrW(2714), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 40: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 41: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 42: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2759), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 43: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2759), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 44: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2753), VBA.ChrW(2690), VBA.ChrW(2734), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 45: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2751), VBA.ChrW(2744), VBA.ChrW(2765), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 46: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715), VBA.ChrW(2759), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 47: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2753), VBA.ChrW(2721), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 48: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2721), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2738), VBA.ChrW(2752), VBA.ChrW(2744)), vbNullString)
            Case 49: LoopkupHString = VBA.Join(Array(VBA.ChrW(2707), VBA.ChrW(2711), VBA.ChrW(2723), VBA.ChrW(2730), VBA.ChrW(2714), VBA.ChrW(2750), VBA.ChrW(2744)), vbNullString)
            Case 50: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2714), VBA.ChrW(2750), VBA.ChrW(2744)), vbNullString)
            Case 51: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709), VBA.ChrW(2750), VBA.ChrW(2741), VBA.ChrW(2728)), vbNullString)
            Case 52: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2750), VBA.ChrW(2741), VBA.ChrW(2728)), vbNullString)
            Case 53: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2759), VBA.ChrW(2730), VBA.ChrW(2728)), vbNullString)
            Case 54: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2763), VBA.ChrW(2730), VBA.ChrW(2728)), vbNullString)
            Case 55: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2690), VBA.ChrW(2714), VBA.ChrW(2750), VBA.ChrW(2741), VBA.ChrW(2728)), vbNullString)
            Case 56: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715), VBA.ChrW(2730), VBA.ChrW(2765), VBA.ChrW(2730), VBA.ChrW(2728)), vbNullString)
            Case 57: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2741), VBA.ChrW(2728)), vbNullString)
            Case 58: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2720), VBA.ChrW(2765), VBA.ChrW(2720), VBA.ChrW(2750), VBA.ChrW(2741), VBA.ChrW(2728)), vbNullString)
            Case 59: LoopkupHString = VBA.Join(Array(VBA.ChrW(2707), VBA.ChrW(2711), VBA.ChrW(2723), VBA.ChrW(2744), VBA.ChrW(2750), VBA.ChrW(2720)), vbNullString)
            Case 60: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2750), VBA.ChrW(2696), VBA.ChrW(2720)), vbNullString)
            Case 61: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709), VBA.ChrW(2744), VBA.ChrW(2720)), vbNullString)
            Case 62: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2720)), vbNullString)
            Case 63: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2759), VBA.ChrW(2744), VBA.ChrW(2720)), vbNullString)
            Case 64: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2763), VBA.ChrW(2744), VBA.ChrW(2720)), vbNullString)
            Case 65: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2750), VBA.ChrW(2690), VBA.ChrW(2744), VBA.ChrW(2720)), vbNullString)
            Case 66: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2720)), vbNullString)
            Case 67: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2721), VBA.ChrW(2744), VBA.ChrW(2720)), vbNullString)
            Case 68: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2721), VBA.ChrW(2744), VBA.ChrW(2720)), vbNullString)
            Case 69: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2711), VBA.ChrW(2723), VBA.ChrW(2763), VBA.ChrW(2744), VBA.ChrW(2751), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 70: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2751), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 71: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 72: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 73: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 74: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2753), VBA.ChrW(2734), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 75: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2690), VBA.ChrW(2714), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 76: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 77: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2751), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 78: LoopkupHString = VBA.Join(Array(VBA.ChrW(2695), VBA.ChrW(2720), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2763), VBA.ChrW(2724), VBA.ChrW(2759), VBA.ChrW(2736)), vbNullString)
            Case 79: LoopkupHString = VBA.Join(Array(VBA.ChrW(2707), VBA.ChrW(2711), VBA.ChrW(2723), VBA.ChrW(2750), VBA.ChrW(2703), VBA.ChrW(2690), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 80: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2690), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 81: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 82: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 83: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 84: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2763), VBA.ChrW(2736), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 85: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2690), VBA.ChrW(2714), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 86: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 87: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2751), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 88: LoopkupHString = VBA.Join(Array(VBA.ChrW(2696), VBA.ChrW(2720), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 89: LoopkupHString = VBA.Join(Array(VBA.ChrW(2728), VBA.ChrW(2759), VBA.ChrW(2741), VBA.ChrW(2765), VBA.ChrW(2735), VBA.ChrW(2750), VBA.ChrW(2744), VBA.ChrW(2752)), vbNullString)
            Case 90: LoopkupHString = VBA.Join(Array(VBA.ChrW(2728), VBA.ChrW(2759), VBA.ChrW(2741), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 91: LoopkupHString = VBA.Join(Array(VBA.ChrW(2703), VBA.ChrW(2709), VBA.ChrW(2750), VBA.ChrW(2723), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 92: LoopkupHString = VBA.Join(Array(VBA.ChrW(2732), VBA.ChrW(2750), VBA.ChrW(2723), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 93: LoopkupHString = VBA.Join(Array(VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2736), VBA.ChrW(2750), VBA.ChrW(2723), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 94: LoopkupHString = VBA.Join(Array(VBA.ChrW(2714), VBA.ChrW(2763), VBA.ChrW(2736), VBA.ChrW(2750), VBA.ChrW(2723), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 95: LoopkupHString = VBA.Join(Array(VBA.ChrW(2730), VBA.ChrW(2690), VBA.ChrW(2714), VBA.ChrW(2750), VBA.ChrW(2723), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 96: LoopkupHString = VBA.Join(Array(VBA.ChrW(2715), VBA.ChrW(2728), VBA.ChrW(2765), VBA.ChrW(2728), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 97: LoopkupHString = VBA.Join(Array(VBA.ChrW(2744), VBA.ChrW(2724), VBA.ChrW(2765), VBA.ChrW(2724), VBA.ChrW(2750), VBA.ChrW(2723), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 98: LoopkupHString = VBA.Join(Array(VBA.ChrW(2693), VBA.ChrW(2720), VBA.ChrW(2765), VBA.ChrW(2720), VBA.ChrW(2750), VBA.ChrW(2723), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
            Case 99: LoopkupHString = VBA.Join(Array(VBA.ChrW(2728), VBA.ChrW(2741), VBA.ChrW(2765), VBA.ChrW(2741), VBA.ChrW(2750), VBA.ChrW(2723), VBA.ChrW(2753), VBA.ChrW(2690)), vbNullString)
        End Select
    End If
    
End Function
