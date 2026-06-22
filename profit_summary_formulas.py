# Auto-generated from Profit Summary.xlsx after row 14 Gaco E5320 insertion and lower-row cleanup.
# One constant per formula cell, plus an aggregate dict.

PS_F_M3 = '=IF(H1="Roofer",\nIF(E5=Data!B4,Data!F4,\nIF(E5=Data!B5,Data!F5,\nIF(E5=Data!B6,Data!F6,\nIF(E5=Data!B7,Data!F7,\nIF(E5=Data!B8,Data!F8,\nIF(E5=Data!B9,Data!F9)))))),\nIF(H1="PCS Direct",\nIF(E5=Data!B13,Data!F13,\nIF(E5=Data!B14,Data!F14,\nIF(E5=Data!B15,Data!F15,\nIF(E5=Data!B16,Data!F16,\nIF(E5=Data!B17,Data!F17,\nIF(E5=Data!B18,Data!F18)))))),""))'
PS_F_P3 = '=IF(M3<>"", ($E3*M3)+E24+E23+E22+E33, "")'
PS_F_M5 = '=IF(M3<>"",M3+40,"")'
PS_F_P5 = '=IF(M5<>"", ($E3*M5)+K24+K23+K22+K33, "")'
PS_F_E7 = '=IF(E3<>"",IF(OR(E5="Ballasted 60 mil", E5="Ballasted 45 mil"), ROUNDUP(E3/30,0), IF(E5="Rock/Foam/Coat", ROUNDUP(E3/75,0), ROUNDUP(E3/45,0))),"")'
PS_F_M7 = '=IF(M3<>"",M3+80,"")'
PS_F_P7 = '=IF(M7<>"", ($E3*M7)+P24+P23+P22+P33, "")'
PS_F_E8 = '=IF(E4<>"",ROUNDUP(E4/45,0),"")'
PS_F_C11 = '=IF(H3="Gaco",ROUNDUP(E3/5*_xlfn.XLOOKUP(E5,Data!B4:B9,Data!C4:C9),0),\nIF(H3="Uniflex",ROUNDUP(E3/5*_xlfn.XLOOKUP(E5,Data!B13:B18,Data!C13:C18),0),""))'
PS_F_D11 = '=IF(C11<>"",IF(H3="Gaco",Data!K8,IF(H3="Uniflex",Data!N8,0)),0)'
PS_F_E11 = '=IF(C11<>"",C11*D11,0)'
PS_F_H11 = '=IF(H3="Gaco",ROUNDUP(E3/5*_xlfn.XLOOKUP(E5,Data!B4:B9,Data!D4:D9),0),\nIF(H3="Uniflex",ROUNDUP(E3/5*_xlfn.XLOOKUP(E5,Data!B13:B18,Data!D13:D18),0),""))'
PS_F_I11 = '=D11'
PS_F_K11 = '=IF(H11<>"",H11*I11,0)'
PS_F_N11 = '=IF(H3="Gaco",ROUNDUP(E3/5*_xlfn.XLOOKUP(E5,Data!B4:B9,Data!E4:E9),0),\nIF(H3="Uniflex",ROUNDUP(E3/5*_xlfn.XLOOKUP(E5,Data!B13:B18,Data!E13:E18),0),""))'
PS_F_O11 = '=D11'
PS_F_P11 = '=IF(N11<>"",N11*O11,0)'
PS_F_S11 = '=IF(T5="10 Year",C11,\nIF(T5= "15 Year",H11,\nIF(T5="20 Year",N11,"")))'
PS_F_T11 = '=IF(S11<>"",D11,0)'
PS_F_U11 = '=IF(S11<>"",S11*T11,0)'
PS_F_C12 = '=IF(H3="Gaco",IF(E5="Rock/Foam/Coat", ROUNDUP(E3*0.03,0), ROUNDUP(E3/10, 0)),"")'
PS_F_D12 = '=IF(C12<>"",Data!K9,0)'
PS_F_E12 = '=IF(C12<>"",C12*D12,0)'
PS_F_H12 = '=C12'
PS_F_I12 = '=D12'
PS_F_K12 = '=IF(H12<>"",H12*I12,0)'
PS_F_N12 = '=C12'
PS_F_O12 = '=D12'
PS_F_P12 = '=IF(N12<>"",N12*O12,0)'
PS_F_S12 = '=IF(T5<>"",C12,"")'
PS_F_T12 = '=IF(S12<>"",D12,0)'
PS_F_U12 = '=IF(S12<>"",S12*T12,0)'
PS_F_C13 = '=IF(AND(H3="Gaco",E5="Mod Bit"),ROUNDUP(E3/5,0),"")'
PS_F_D13 = '=IF(C13<>"",Data!K11,0)'
PS_F_E13 = '=IF(C13<>"",C13*D13,0)'
PS_F_H13 = '=C13'
PS_F_I13 = '=D13'
PS_F_K13 = '=IF(H13<>"",H13*I13,0)'
PS_F_N13 = '=C13'
PS_F_O13 = '=D13'
PS_F_P13 = '=IF(N13<>"",N13*O13,0)'
PS_F_S13 = '=IF(T5<>"",C13,"")'
PS_F_T13 = '=IF(S13<>"",D13,0)'
PS_F_U13 = '=IF(S13<>"",S13*T13,0)'
PS_F_C14 = '=""'
PS_F_D14 = '=IF(C14<>"",Data!K10,0)'
PS_F_E14 = '=IF(C14<>"",C14*D14,0)'
PS_F_H14 = '=C14'
PS_F_I14 = '=D14'
PS_F_K14 = '=IF(H14<>"",H14*I14,0)'
PS_F_N14 = '=C14'
PS_F_O14 = '=D14'
PS_F_P14 = '=IF(N14<>"",N14*O14,0)'
PS_F_S14 = '=IF(T5<>"",C14,"")'
PS_F_T14 = '=IF(S14<>"",D14,0)'
PS_F_U14 = '=IF(S14<>"",S14*T14,0)'
PS_F_C15 = '=IF(AND(H3="Uniflex",E5="TPO/EPDM"),ROUNDUP(E3/20,0),\nIF(AND(H3="Uniflex",E5="Metal"),ROUNDUP(E3/10,0),\nIF(AND(H3="Uniflex",E5="Mod Bit"),ROUNDUP(E3/20,0),\nIF(AND(H3="Uniflex",E5="Ballasted 60 mil"),ROUNDUP(E3/10,0),\nIF(AND(H3="Uniflex",E5="Ballasted 45 mil"),ROUNDUP(E3/10,0),\nIF(AND(H3="Uniflex",E5="Rock/Foam/Coat"),ROUNDUP(E3/20,0),""))))))'
PS_F_D15 = '=IF(C15<>"",Data!N9,0)'
PS_F_E15 = '=IF(C15<>"",C15*D15,0)'
PS_F_H15 = '=C15'
PS_F_I15 = '=D15'
PS_F_K15 = '=IF(H15<>"",H15*I15,0)'
PS_F_N15 = '=C15'
PS_F_O15 = '=D15'
PS_F_P15 = '=IF(N15<>"",N15*O15,0)'
PS_F_S15 = '=IF(T5<>"",C15,"")'
PS_F_T15 = '=IF(S15<>"",D15,0)'
PS_F_U15 = '=IF(S15<>"",S15*T15,0)'
PS_F_C16 = '=IF(AND(H3="Uniflex",E5="Mod Bit"),ROUNDUP(E3/5,0),"")'
PS_F_D16 = '=IF(C16<>"",Data!N10,0)'
PS_F_E16 = '=IF(C16<>"",C16*D16,0)'
PS_F_H16 = '=C16'
PS_F_I16 = '=D16'
PS_F_K16 = '=IF(H16<>"",H16*I16,0)'
PS_F_N16 = '=C16'
PS_F_O16 = '=D16'
PS_F_P16 = '=IF(N16<>"",N16*O16,0)'
PS_F_S16 = '=IF(T5<>"",C16,"")'
PS_F_T16 = '=IF(S16<>"",D16,0)'
PS_F_U16 = '=IF(S16<>"",S16*T16,0)'
PS_F_C17 = '=IF(AND(OR(E5="Ballasted 60 mil",E5="Ballasted 45 mil"),E3<>""),ROUNDUP(E3/18,0),"")'
PS_F_D17 = '=IF(C17<>"",Data!K12,0)'
PS_F_E17 = '=IF(C17<>"",C17*D17,0)'
PS_F_H17 = '=C17'
PS_F_I17 = '=D17'
PS_F_K17 = '=IF(H17<>"",H17*I17,0)'
PS_F_N17 = '=C17'
PS_F_O17 = '=D17'
PS_F_P17 = '=IF(N17<>"",N17*O17,0)'
PS_F_S17 = '=IF(T5<>"",C17,"")'
PS_F_T17 = '=IF(S17<>"",D17,0)'
PS_F_U17 = '=IF(S17<>"",S17*T17,0)'
PS_F_C18 = '=IF(AND(E3<>"",E5="Rock/Foam/Coat"),ROUNDUP(E3/25,0),"")'
PS_F_D18 = '=IF(AND(E5="Rock/Foam/Coat",H3="Gaco"),Data!N15,\nIF(AND(E5="Rock/Foam/Coat",H3="Uniflex"), Data!N16,\n0))'
PS_F_E18 = '=IF(E5="Rock/Foam/Coat",C18*D18,0)'
PS_F_H18 = '=C18'
PS_F_I18 = '=D18'
PS_F_K18 = '=IF(E5="Rock/Foam/Coat",H18*I18,0)'
PS_F_N18 = '=C18'
PS_F_O18 = '=D18'
PS_F_P18 = '=IF(E5="Rock/Foam/Coat",N18*O18,0)'
PS_F_S18 = '=IF(T5<>"",C18,"")'
PS_F_T18 = '=IF(T5<>"",D18,0)'
PS_F_U18 = '=IF(S18<>"",S18*T18,0)'
PS_F_D19 = '=IF(E5="Rock/Foam/Coat",Data!N18,0)'
PS_F_E19 = '=IF(AND(E3<>"",D19<>0), E3*D19,0)'
PS_F_I19 = '=D19'
PS_F_K19 = '=E19'
PS_F_O19 = '=D19'
PS_F_P19 = '=E19'
PS_F_T19 = '=IF(T5<>"",D19,0)'
PS_F_U19 = '=IF(T5<>"",E19,0)'
PS_F_K20 = '=E20'
PS_F_P20 = '=E20'
PS_F_U20 = '=IF(T5<>"",E20,0)'
PS_F_C21 = '=IF(E7<>"", E7,"")'
PS_F_D21 = '=IF(C21<>"",Data!J3,0)'
PS_F_E21 = '=IF(C21<>"",H21*D21,0)'
PS_F_H21 = '=C21'
PS_F_I21 = '=D21'
PS_F_K21 = '=IF(H21<>"",H21*I21,0)'
PS_F_N21 = '=C21'
PS_F_O21 = '=D21'
PS_F_P21 = '=IF(N21<>"",N21*O21,0)'
PS_F_S21 = '=IF(T5<>"",C21,"")'
PS_F_T21 = '=IF(T5<>"",D21,0)'
PS_F_U21 = '=IF(S21<>"",S21*T21,0)'
PS_F_K22 = '=E22'
PS_F_P22 = '=E22'
PS_F_U22 = '=IF(T5<>"",E22,0)'
PS_F_K23 = '=E23'
PS_F_P23 = '=E23'
PS_F_U23 = '=IF(T5<>"",E23,0)'
PS_F_E24 = '=IF(AND(H3="Gaco",H5="Yes"),500,0)'
PS_F_K24 = '=IF(AND(H3="Gaco",H5="Yes"),500,0)'
PS_F_P24 = '=IF(AND(H3="Gaco",H5="Yes"),500,0)'
PS_F_U24 = '=IF(T5="10 Year",E24,\nIF(T5="15 Year",K24,\nIF(T5="20 Year",P24,0)))'
PS_F_E25 = '=IF(P3="",0,IF(AND(H7<>"Mark",H7<>"Richard"),ROUND((P3-E18-E19-E20-E22-E23)*Data!K15,0),0))'
PS_F_K25 = '=IF(P5="",0,IF(AND(H7<>"Mark",H7<>"Richard"),ROUND((P5-K18-K19-K20-K22-K23)*Data!K15,0),0))'
PS_F_P25 = '=IF(P7="",0,IF(AND(H7<>"Mark",H7<>"Richard"),ROUND((P7-P18-P19-P20-P22-P23)*Data!K15,0),0))'
PS_F_U25 = '=IF(T5="10 Year",E25,\nIF(T5="15 Year",K25,\nIF(T5="20 Year",P25,0)))'
PS_F_E27 = '=IF(E11<>0,SUM(E11:E25),0)'
PS_F_K27 = '=IF(K11<>0,SUM(K11:K25),0)'
PS_F_P27 = '=IF(P11<>0,SUM(P11:P25),0)'
PS_F_U27 = '=IF(U11<>0,SUM(U11:U25),0)'
PS_F_E29 = '=IF(P3<>"",ROUND(P3-E27-E32,0),0)'
PS_F_K29 = '=IF(P5<>"",P5-K27-K32,0)'
PS_F_P29 = '=IF(P7<>"",P7-P27-P32,0)'
PS_F_U29 = '=IF(T5="10 Year",P3-U27-U32,\nIF(T5="15 Year",P5-U27-U32,\nIF(T5="20 Year",P7-U27-U32,0)))'
PS_F_E30 = '=IF(AND(E29<>0,P3<>0),ROUND(E29/P3,2),"")'
PS_F_K30 = '=IF(AND(K29<>0,P5<>0),K29/P5,"")'
PS_F_P30 = '=IF(AND(P29<>0,P7<>0),P29/P7,"")'
PS_F_U30 = '=IF(T5="10 Year",U29/P3,\nIF(T5="15 Year",U29/P5,\nIF(T5="20 Year",U29/P7,"")))'
PS_F_E31 = '=IF(E29<>0,ROUND(E29/C21,0),0)'
PS_F_K31 = '=IF(K29<>0,ROUND(K29/H21,0),0)'
PS_F_P31 = '=IF(P29<>0,ROUND(P29/N21,0),0)'
PS_F_U31 = '=IF(U29<>0,U29/S21,0)'
PS_F_E32 = '=IF(E27<>0,ROUND(Data!K18*(P3-E27),0),0)'
PS_F_K32 = '=IF(K27<>0,ROUND(Data!K18*(P5-K27),0),0)'
PS_F_P32 = '=IF(P27<>0,ROUND(Data!K18*(P7-P27),0),0)'
PS_F_U32 = '=IF(T5="10 Year",ROUND(Data!K18*(P3-U27),0),\nIF(T5="15 Year",ROUND(Data!K18*(P5-U27),0),\nIF(T5="20 Year",ROUND(Data!K18*(P7-U27),0),0)))'
PS_F_E33 = '= IF(M3<>"", IF(H7="Mark", ROUND((($E3*M3)+E24+E23+E22)*Data!N4,0), ROUND((($E3*M3)+E24+E23+E22)*Data!N3,0)),0)'
PS_F_K33 = '= IF(M5<>"", IF(H7="Mark", ROUND((($E3*M5)+K24+K23+K22)*Data!N4,0), ROUND((($E3*M5)+K24+K23+K22)*Data!N3,0)),0)'
PS_F_P33 = '= IF(M7<>"", IF(H7="Mark", ROUND((($E3*M7)+P24+P23+P22)*Data!N4,0), ROUND((($E3*M7)+P24+P23+P22)*Data!N3,0)),0)'
PS_F_U33 = '=IF(T5="10 Year",E33,\nIF(T5="15 Year",K33,\nIF(T5="20 Year",P33,0)))'

PROFIT_SUMMARY_FORMULAS = {
    'M3': PS_F_M3,
    'P3': PS_F_P3,
    'M5': PS_F_M5,
    'P5': PS_F_P5,
    'E7': PS_F_E7,
    'M7': PS_F_M7,
    'P7': PS_F_P7,
    'E8': PS_F_E8,
    'C11': PS_F_C11,
    'D11': PS_F_D11,
    'E11': PS_F_E11,
    'H11': PS_F_H11,
    'I11': PS_F_I11,
    'K11': PS_F_K11,
    'N11': PS_F_N11,
    'O11': PS_F_O11,
    'P11': PS_F_P11,
    'S11': PS_F_S11,
    'T11': PS_F_T11,
    'U11': PS_F_U11,
    'C12': PS_F_C12,
    'D12': PS_F_D12,
    'E12': PS_F_E12,
    'H12': PS_F_H12,
    'I12': PS_F_I12,
    'K12': PS_F_K12,
    'N12': PS_F_N12,
    'O12': PS_F_O12,
    'P12': PS_F_P12,
    'S12': PS_F_S12,
    'T12': PS_F_T12,
    'U12': PS_F_U12,
    'C13': PS_F_C13,
    'D13': PS_F_D13,
    'E13': PS_F_E13,
    'H13': PS_F_H13,
    'I13': PS_F_I13,
    'K13': PS_F_K13,
    'N13': PS_F_N13,
    'O13': PS_F_O13,
    'P13': PS_F_P13,
    'S13': PS_F_S13,
    'T13': PS_F_T13,
    'U13': PS_F_U13,
    'C14': PS_F_C14,
    'D14': PS_F_D14,
    'E14': PS_F_E14,
    'H14': PS_F_H14,
    'I14': PS_F_I14,
    'K14': PS_F_K14,
    'N14': PS_F_N14,
    'O14': PS_F_O14,
    'P14': PS_F_P14,
    'S14': PS_F_S14,
    'T14': PS_F_T14,
    'U14': PS_F_U14,
    'C15': PS_F_C15,
    'D15': PS_F_D15,
    'E15': PS_F_E15,
    'H15': PS_F_H15,
    'I15': PS_F_I15,
    'K15': PS_F_K15,
    'N15': PS_F_N15,
    'O15': PS_F_O15,
    'P15': PS_F_P15,
    'S15': PS_F_S15,
    'T15': PS_F_T15,
    'U15': PS_F_U15,
    'C16': PS_F_C16,
    'D16': PS_F_D16,
    'E16': PS_F_E16,
    'H16': PS_F_H16,
    'I16': PS_F_I16,
    'K16': PS_F_K16,
    'N16': PS_F_N16,
    'O16': PS_F_O16,
    'P16': PS_F_P16,
    'S16': PS_F_S16,
    'T16': PS_F_T16,
    'U16': PS_F_U16,
    'C17': PS_F_C17,
    'D17': PS_F_D17,
    'E17': PS_F_E17,
    'H17': PS_F_H17,
    'I17': PS_F_I17,
    'K17': PS_F_K17,
    'N17': PS_F_N17,
    'O17': PS_F_O17,
    'P17': PS_F_P17,
    'S17': PS_F_S17,
    'T17': PS_F_T17,
    'U17': PS_F_U17,
    'C18': PS_F_C18,
    'D18': PS_F_D18,
    'E18': PS_F_E18,
    'H18': PS_F_H18,
    'I18': PS_F_I18,
    'K18': PS_F_K18,
    'N18': PS_F_N18,
    'O18': PS_F_O18,
    'P18': PS_F_P18,
    'S18': PS_F_S18,
    'T18': PS_F_T18,
    'U18': PS_F_U18,
    'D19': PS_F_D19,
    'E19': PS_F_E19,
    'I19': PS_F_I19,
    'K19': PS_F_K19,
    'O19': PS_F_O19,
    'P19': PS_F_P19,
    'T19': PS_F_T19,
    'U19': PS_F_U19,
    'K20': PS_F_K20,
    'P20': PS_F_P20,
    'U20': PS_F_U20,
    'C21': PS_F_C21,
    'D21': PS_F_D21,
    'E21': PS_F_E21,
    'H21': PS_F_H21,
    'I21': PS_F_I21,
    'K21': PS_F_K21,
    'N21': PS_F_N21,
    'O21': PS_F_O21,
    'P21': PS_F_P21,
    'S21': PS_F_S21,
    'T21': PS_F_T21,
    'U21': PS_F_U21,
    'K22': PS_F_K22,
    'P22': PS_F_P22,
    'U22': PS_F_U22,
    'K23': PS_F_K23,
    'P23': PS_F_P23,
    'U23': PS_F_U23,
    'E24': PS_F_E24,
    'K24': PS_F_K24,
    'P24': PS_F_P24,
    'U24': PS_F_U24,
    'E25': PS_F_E25,
    'K25': PS_F_K25,
    'P25': PS_F_P25,
    'U25': PS_F_U25,
    'E27': PS_F_E27,
    'K27': PS_F_K27,
    'P27': PS_F_P27,
    'U27': PS_F_U27,
    'E29': PS_F_E29,
    'K29': PS_F_K29,
    'P29': PS_F_P29,
    'U29': PS_F_U29,
    'E30': PS_F_E30,
    'K30': PS_F_K30,
    'P30': PS_F_P30,
    'U30': PS_F_U30,
    'E31': PS_F_E31,
    'K31': PS_F_K31,
    'P31': PS_F_P31,
    'U31': PS_F_U31,
    'E32': PS_F_E32,
    'K32': PS_F_K32,
    'P32': PS_F_P32,
    'U32': PS_F_U32,
    'E33': PS_F_E33,
    'K33': PS_F_K33,
    'P33': PS_F_P33,
    'U33': PS_F_U33,
}
