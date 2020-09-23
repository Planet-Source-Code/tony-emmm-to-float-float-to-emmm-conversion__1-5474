Attribute VB_Name = "modEMMM"
Option Explicit

'I have no idea if the math still works if either of the
'following two constants is changed! ;-)
Public Const TotalBits = 16
Public Const MantissaBits = 12 'must be less than TotalBits

Public Const MaxPower = 2 ^ (TotalBits - MantissaBits) - 1
Public Const MaxM = 2 ^ MantissaBits 'Maximum Mantissa Value +1

Public Const MinEMMM = 1
Public Const MaxEMMM = 65528 'FloatFromEMMM(MaxPower, MaxM - 1)

Public Function Log2(x As Variant)
  Log2 = Log(x) / Log(2)
End Function

Public Function Pwr2(ByVal Power As Integer) As Long
  Static Powers(0 To MaxPower) As Long
  
  Dim i As Integer
  
  If Powers(0) = 0 Then 'powers not yet initialized
    For i = 0 To MaxPower
      Powers(i) = 2 ^ i
    Next i
  End If 'powers not yet initialized
  
  If Power >= 0 And Power <= MaxPower Then Pwr2 = Powers(Power)
End Function

Public Function FloatFromEMMM(ByVal e As Byte, ByVal mmm As Integer) As Single
  If e > 15 Or mmm < 0 Or mmm >= MaxM Then Exit Function 'bad input
  
  FloatFromEMMM = mmm * 2 ^ (e - MantissaBits) + 2 ^ e
  
End Function

Public Sub EMMMFromFloat(ByVal f As Single, ByRef e As Byte, ByRef mmm As Integer)
  If f < MinEMMM Or f > MaxEMMM Then Exit Sub 'bad input
    
  'given f, solve for mmm as a function of f and E:
  '   mmm = maxm * (f/2^E - 1)               [1]
  'There would be infinite solutions except that we require e and mmm to be integers, and also:
  '   0 <= mmm < maxm
  'or
  '   0 <= maxm * (f/2^E - 1) < maxm
  'which eventually leads to
  '   log2(f/2) < E <= log2(f)
  'This yields a unique value for E and then mmm can be calculated using equation [1].
  
  e = Int(Log2(f))
  mmm = Int(MaxM * (f / Pwr2(e) - 1)) 'originally did not have the int( ... )
  
  'the following code was removed in favor of the int( ... ).  The float in question
  'was 32766, which can be e=14,mmm=4095, or e=15,mmm=0.  Using the current code, it
  'comes back as 32764, using the code below it comes back as 32768.  Take your pick.
  '
  'If mmm >= MaxM Then
  '  'log2(f) was X.9999etc, with 9s going for more than there are significant digits
  '  'so int() was a little harsh - int(log2(f)+0.0001) would work but the fuzz would
  '  'depend on constants (mantissabits etc).  Hopefully this check is more robust.
  '  'But since I didn't expect this problem I can't be sure. Maybe another solution
  '  'would be to calc mmm = int( ... ) since the implicit conversion rounds.
  '  e = e + 1
  '  mmm = 0
  'End If
  
End Sub

Public Sub test(Form1 As Form1)
On Error Resume Next
  Const SignificantDigits = 4 'this is of course somehow related to the EMMM constants!

  Dim e1 As Byte, e2 As Byte
  Dim m1 As Integer, m2 As Integer
  Dim f1 As Single, f2 As Single
  Dim f1comp As Long, f2comp As Long
  
  Dim i As Long
  Dim fmt As String
  
  Debug.Print "MaxEMMM = "; FloatFromEMMM(MaxPower, MaxM - 1)
  Randomize
  fmt = "0." & String$(SignificantDigits, "0") & "e-00"
  
  With Form1.picemmm
    .ScaleHeight = MaxPower + 1
    .ScaleWidth = MaxM
  End With
  
  With Form1.picf
    .ScaleHeight = MaxEMMM
    .ScaleWidth = 1
  End With
  
  Do
    'calc float from emmm
    e1 = Int(Rnd * 16)
    m1 = Int(Rnd * MaxM)
    f1 = FloatFromEMMM(e1, m1)
    
    'check it
    EMMMFromFloat f1, e2, m2
    If e1 <> e2 Or m1 <> m2 Then Exit Sub 'Stop

Debug.Print (m1 + (e1 * 4096)), f1, (m2 + (e2 * 4096))

    'report it OK
    Form1.picemmm.Line (m1, e1)-(m1 + 1, e1 + 1), , BF
    
    'calc emmm from float
    f1 = MinEMMM + Rnd * (MaxEMMM - MinEMMM)
    EMMMFromFloat f1, e1, m1
    
    'check it
    f2 = FloatFromEMMM(e1, m1)
    f1comp = CSng(Left$(Format(f1, fmt), SignificantDigits + 2)) * 10 ^ SignificantDigits
    f2comp = CSng(Left$(Format(f2, fmt), SignificantDigits + 2)) * 10 ^ SignificantDigits
    If Abs(f1comp - f2comp) > 10 Then
      Select Case CInt(Right$(Format(f1, fmt), 2)) - CInt(Right$(Format(f2, fmt), 2))
        Case -1
          If Abs((f1comp \ 10) - f2comp) > 10 Then Stop
        Case 1
          If Abs(f1comp - (f2comp \ 10)) > 10 Then Stop
        Case Else
Exit Sub
'          Stop
      End Select
    End If
    
    'report it OK
    Form1.picf.Line (0, f1)-(1, f1)
    
    'increment/report counter
    i = i + 1
    If (i Mod 1000) = 0 Then Debug.Print Now; ": "; i
    DoEvents
  Loop While i < 2 ^ 31 And Not Form1.flgStop
End Sub
