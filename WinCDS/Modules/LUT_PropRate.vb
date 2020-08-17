Imports Microsoft.VisualBasic.Interaction
Public Module LUT_PropRate
    Public Enum PropInsuranceTypes
        SingleInterest
        DualInterest
    End Enum

    Public ReadOnly Property propInsurance(ByVal rateType As PropInsuranceTypes) As Single
        Get
            On Error Resume Next
            propInsurance = 0
            propInsurance = Switch(
                            (rateType = PropInsuranceTypes.SingleInterest), 0.87 _
                          , (rateType = PropInsuranceTypes.DualInterest), 1.31
                        )
        End Get
    End Property
End Module
