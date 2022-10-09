' Выцепляем число скважин из GAP
WellNum = DoGet("GAP.MOD[{PROD}].Well.Count")
j = 1

'logmsg(PG_rate)
Dim Arr()

PG_rate = DoGet("GAP.MOD[{PROD}].GROUP[{Group1}].SolverResults[0].QgasAvg")
Rate = DoGet("GAP.MOD[{PROD}].SEP[{CPF_Inlet}].SolverResults[0].QgasAvg")

Do While ((Rate < 29500 Or Rate > 30500) And j < 4) Or ((PG_rate < 15000 Or PG_rate > 15900) And j < 4)
    'PG_rate = DoGet("GAP.MOD[{PROD}].SEP[{SeparatorA}].SolverResults[0].GasRate")
    ReDim Arr(0)
    Gdpmin = 100
    Gdpmax = 0
    Z = 0
    LogMsg ("PG_Rate_before = " + CStr(PG_rate))
    LogMsg ("PNG_rate_before = " + CStr(Rate))
    ''Get GdpMin and GdpMax
    For i = 0 To (WellNum - 1)
        L = DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].Label")
        L1 = Left(L, 1)
        grat = CDbl(DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].SolverResults[0].GasRate"))
        If grat <> "FNA" And L1 = "3" Then
            Gdp = CDbl(DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue"))
            ReDim Preserve Arr(Z)
            'logmsg(Arr(GDp))
            Arr(Z) = Gdp
            'logmsg("GDp = " + Cstr(GDp))
            Z = Z + 1
        End If
    Next
    For Each x In Arr
        If x > Gdpmax Then
            Gdpmax = x
        End If
        If x < Gdpmin Then
            Gdpmin = x
        End If
    Next
    
    j = j + 1
    'logmsg("j=" + cstr(j))
    
    ''Work with Oil and gas wells
    For i = 0 To (WellNum - 1)
        
        grat = CDbl(DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].SolverResults[0].GasRate"))
        L = DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].Label")
        Length = 1
        L1 = Left(L, Length)
        
        'PG_rate = DoGet("GAP.MOD[{PROD}].SEP[{Separator}].SolverResults[0].GasRate")
        
        If grat <> "FNA" And L1 <> "3" And (Rate < 29500 Or Rate > 30500) Then
            If grat >= 1500 Then
                DoCmd ("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].MASK")
            ElseIf grat < 1500 Then
                Dp = CDbl(DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue"))
                If Rate < 29500 Then
                    If grat > 1 Then
                        x = 2.39 * grat^(-0.12) + 0.014
                        newDP = Dp / x
                        DoSet "GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue", newDP
                        LogMsg ("Well " + CStr(L) + "Change_DP  " + CStr(Dp) + "on   " + CStr(newDP))
                    End If
                ElseIf Rate > 30500 Then
                    If grat > 1 Then
                        If Dp <= 10 And grat > 100 Then
                            newDP = Dp + 10
                            DoSet "GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue", newDP
                            LogMsg (CStr(L) + "Change_DP  " + CStr(Dp) + "on   " + CStr(newDP))
                        ElseIf Dp > 10 Then
                            x = 3.22 * grat^(-0.25)
                            newDP = Dp / x
                            If newDP > 200 Then
                                newDP = 200
                            End If
                            DoSet "GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue", newDP
                            LogMsg (CStr(L) + "Change_DP  " + CStr(Dp) + "on   " + CStr(newDP))
                        End If
                    End If
                End If
            End If
        End If
        
        'LogMsg("PGR = " + cstr(PG_rate)+"  GRAT =  "+ cstr(grat)+"  Wname = "+ cstr(L))
        grat = CDbl(DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].SolverResults[0].GasRate"))
        L = DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].Label")
        L1 = Left(L, 1)
        'PG_rate = DoGet("GAP.MOD[{PROD}].SEP[{Separator}].SolverResults[0].GasRate")
        PG_rate = DoGet("GAP.MOD[{PROD}].GROUP[{Group1}].SolverResults[0].QgasAvg")
        Rate = DoGet("GAP.MOD[{PROD}].SEP[{CPF_Inlet}].SolverResults[0].QgasAvg")
        
        If grat <> "FNA" And L1 = "3" And PG_rate < 15000 Then
            Gdp = CDbl(DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue"))
            If Gdp = Gdpmax Then
                newDP = Gdp * 0.9
                DoSet "GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue", newDP
                LogMsg (CStr(L) + "__Change_DP__" + CStr(Gdp) + "__On__" + CStr(newDP))
            End If
        End If
        
        If grat <> "FNA" And L1 = "3" And PG_rate > 15900 Then
            Gdp = CDbl(DoGet("GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue"))
            If Gdp = Gdpmin Then
                newDP = (Gdp + 10) * 1.1
                DoSet "GAP.MOD[{PROD}].WELL[" + CStr(i) + "].DPControlValue", newDP
                LogMsg (CStr(L) + "__Change_DP__" + CStr(Gdp) + "__On __" + CStr(newDP))
            End If
        End If
        
    Next
    DoCmd "GAP.SOLVENETWORK(0,MOD[{PROD}])"
    
    PG_rate = DoGet("GAP.MOD[{PROD}].GROUP[{Group1}].SolverResults[0].QgasAvg")
    Rate = DoGet("GAP.MOD[{PROD}].SEP[{CPF_Inlet}].SolverResults[0].QgasAvg")
    
    LogMsg ("PG_Rate_after = " + CStr(PG_rate))
    LogMsg ("PNG_rate_after = " + CStr(Rate))
    LogMsg ("Solve!")
Loop