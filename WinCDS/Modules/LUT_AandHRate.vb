Imports Microsoft.VisualBasic.Interaction
Public Module LUT_AandHRate
    Public Enum BenefitTypes
        Nonretro_14day
        Nonretro_30day
        Retro_7day
        Retro_14day
        retro_30day
    End Enum

    Public Enum RateTypes
        SingleRate
        JointRate
    End Enum

    Public ReadOnly Property creditAHRate(ByVal Months As Integer, ByVal rateType As RateTypes, ByVal benefitType As BenefitTypes) As Single
        Get
            On Error Resume Next
            creditAHRate = 0
            creditAHRate = Switch(
                  (rateType = RateTypes.SingleRate) _
                  , Switch(
                            (benefitType = BenefitTypes.Nonretro_14day), Choose(Months / 12, 1.4, 1.9, 2.4, 2.85, 3.35, 3.85, 4.3, 4.8, 5.25, 5.75) _
                          , (benefitType = BenefitTypes.Nonretro_30day), Choose(Months / 12, 0.95, 1.4, 1.9, 2.4, 2.85, 3.35, 3.85, 4.3, 4.8, 5.25) _
                          , (benefitType = BenefitTypes.Retro_7day), Choose(Months / 12, 2.6, 3.5, 4.35, 5.25, 6.1) _
                          , (benefitType = BenefitTypes.Retro_14day), Choose(Months / 12, 2.1, 2.85, 3.65, 4.4, 5.2, 5.95, 6.7, 7.5, 8.25, 9) _
                          , (benefitType = BenefitTypes.retro_30day), Choose(Months / 12, 1.4, 1.9, 2.4, 2.85, 3.35, 3.85, 4.3, 4.8, 5.25, 5.75)
                        ) _
                  , (rateType = RateTypes.JointRate) _
                  , Switch(
                          (benefitType = BenefitTypes.Nonretro_14day), Choose(Months / 12, 2.33, 3.16, 3.99, 4.74, 5.58, 6.41, 7.16, 7.99, 8.74, 9.57) _
                        , (benefitType = BenefitTypes.Nonretro_30day), Choose(Months / 12, 1.58, 2.33, 3.16, 3.99, 4.74, 5.58, 6.41, 7.16, 7.99, 8.74) _
                        , (benefitType = BenefitTypes.Retro_7day), Choose(Months / 12, 4.33, 5.83, 7.24, 8.74, 10.16) _
                        , (benefitType = BenefitTypes.Retro_14day), Choose(Months / 12, 3.49, 4.74, 6.08, 7.33, 8.66, 9.91, 11.16, 12.49, 13.74, 14.99) _
                        , (benefitType = BenefitTypes.retro_30day), Choose(Months / 12, 2.33, 3.16, 3.99, 4.74, 5.58, 6.41, 7.16, 7.99, 8.74, 9.57)
                        )
                   )
        End Get
    End Property
End Module
