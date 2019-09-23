Module modAmericanHeritageInsurance
    Public Function AmericanHeritage_Prop(ByVal Amount As Decimal, ByVal Term As Long) As Decimal
        Select Case Term
            '    Case 1: AmericanHeritage_Prop = 0.33
            '    Case 2: AmericanHeritage_Prop = 0.65
            '    Case 3: AmericanHeritage_Prop = 0.98
            '    Case 4: AmericanHeritage_Prop = 1.3
            '    Case 5: AmericanHeritage_Prop = 1.63
            '    Case 6: AmericanHeritage_Prop = 1.95
            '    Case 7: AmericanHeritage_Prop = 2.28
            '    Case 8: AmericanHeritage_Prop = 2.6
            '    Case 9: AmericanHeritage_Prop = 2.93
            '    Case 10: AmericanHeritage_Prop = 3.25
            '    Case 11: AmericanHeritage_Prop = 3.58
            '    Case 12: AmericanHeritage_Prop = 3.9
            '    Case 13: AmericanHeritage_Prop = 4.23
            '    Case 14: AmericanHeritage_Prop = 4.55
            '    Case 15: AmericanHeritage_Prop = 4.88
            '    Case 16: AmericanHeritage_Prop = 5.2
            '    Case 17: AmericanHeritage_Prop = 5.53
            '    Case 18: AmericanHeritage_Prop = 5.85
            '    Case 19: AmericanHeritage_Prop = 6.18
            '    Case 20: AmericanHeritage_Prop = 6.5
            '    Case 21: AmericanHeritage_Prop = 6.83
            '    Case 22: AmericanHeritage_Prop = 7.15
            '    Case 23: AmericanHeritage_Prop = 7.48
            '    Case 24: AmericanHeritage_Prop = 7.8
            '    Case 25: AmericanHeritage_Prop = 8.13
            '    Case 26: AmericanHeritage_Prop = 8.45
            '    Case 27: AmericanHeritage_Prop = 8.78
            '    Case 28: AmericanHeritage_Prop = 9.1
            '    Case 29: AmericanHeritage_Prop = 9.43
            '    Case 30: AmericanHeritage_Prop = 9.75
            '    Case 31: AmericanHeritage_Prop = 10.08
            '    Case 32: AmericanHeritage_Prop = 10.4
            '    Case 33: AmericanHeritage_Prop = 10.73
            '    Case 34: AmericanHeritage_Prop = 11.05
            '    Case 35: AmericanHeritage_Prop = 11.38
            '    Case 36: AmericanHeritage_Prop = 11.7
        End Select

        AmericanHeritage_Prop = AmericanHeritage_Prop * (Amount / 100)
    End Function

End Module
