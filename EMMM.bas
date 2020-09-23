Attribute VB_Name = "mEMMM"
Option Explicit

Private Const TotalBits = 16
Private Const MantissaBits = 12 'must be less than TotalBits

Private Const MaxPower = 2 ^ (TotalBits - MantissaBits) - 1
Private Const MaxM = 2 ^ MantissaBits 'Maximum Mantissa Value +1

Private Const MinEMMM = 1
Private Const MaxEMMM = 65528

Private Function Log2(x As Variant)
    Log2 = Log(x) / Log(2)
End Function

Private Function Pwr2(ByVal Power As Integer) As Long
    Static Powers(0 To MaxPower) As Long
  
    Dim i As Integer
  
    If Powers(0) = 0 Then 'powers not yet initialized
        For i = 0 To MaxPower
            Powers(i) = 2 ^ i
        Next i
    End If
  
    If Power >= 0 And Power <= MaxPower Then Pwr2 = Powers(Power)
End Function

Public Function EMMM2Float(nEMMM As Long) As Single
    Dim e As Byte, mmm As Integer

    e = (nEMMM And 61440) / 4096&
    mmm = nEMMM And 4095&
    'If e > 15 Or mmm < 0 Or mmm >= MaxM Then Exit Function 'bad input

    EMMM2Float = mmm * 2 ^ (e - MantissaBits) + 2 ^ e
End Function

Public Function Float2EMMM(ByVal f As Single) As Long
    If f < MinEMMM Then
        Float2EMMM = 0
        Exit Function
    ElseIf f > MaxEMMM Then
        Float2EMMM = 65535
        Exit Function
    End If
    
    ' Given f, solve for mmm as a function of f and E:
    '   mmm = maxm * (f/2^E - 1)               [1]
    ' There would be infinite solutions except that we require e and mmm to be integers, and also:
    '   0 <= mmm < maxm
    ' or
    '   0 <= maxm * (f/2^E - 1) < maxm
    ' which eventually leads to
    '   log2(f/2) < E <= log2(f)
    ' This yields a unique value for E and then mmm can be calculated using equation [1].
  
    Dim e As Byte, mmm As Integer

    e = Int(Log2(f))
    mmm = Int(MaxM * (f / Pwr2(e) - 1))
    Float2EMMM = mmm + (e * 4096&)
End Function
