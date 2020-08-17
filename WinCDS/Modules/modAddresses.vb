Module modAddresses
    Public Function CleanState(ByVal ST As String) As String
        CleanState = Trim(LCase(ST))

        If Left(CleanState, 6) = "south " Then CleanState = "s " & Mid(CleanState, 7)
        If Left(CleanState, 6) = "north " Then CleanState = "n " & Mid(CleanState, 7)
        If Left(CleanState, 5) = "west " Then CleanState = "w " & Mid(CleanState, 6)
        If Left(CleanState, 4) = "new " Then CleanState = "n " & Mid(CleanState, 5)

        If CleanState Like "alab*" Then CleanState = "AL"
        If CleanState Like "alak*" Then CleanState = "AK"
        If CleanState Like "ari*" Then CleanState = "AZ"
        If CleanState Like "ark*" Then CleanState = "AK"
        If CleanState Like "ca*" Then CleanState = "CA"
        If CleanState Like "col*" Then CleanState = "CO"
        If CleanState Like "con*" Then CleanState = "CT"
        If CleanState Like "de*" Then CleanState = "DE"
        If CleanState Like "di*" Then CleanState = "DC"
        If CleanState Like "f*" Then CleanState = "FL"
        If CleanState Like "g*" Then CleanState = "GA"
        If CleanState Like "h*" Then CleanState = "HI"
        If CleanState Like "id*" Then CleanState = "ID"
        If CleanState Like "il*" Then CleanState = "IL"
        If CleanState Like "in*" Then CleanState = "IN"
        If CleanState Like "io*" Then CleanState = "IA"
        If CleanState Like "ka*" Then CleanState = "KS"
        If CleanState Like "ke*" Then CleanState = "KY"
        If CleanState Like "l*" Then CleanState = "LA"
        If CleanState Like "mai*" Then CleanState = "ME"
        If CleanState Like "mar*" Then CleanState = "MD"
        If CleanState Like "mas*" Then CleanState = "MA"
        If CleanState Like "mic*" Then CleanState = "MI"
        If CleanState Like "min*" Then CleanState = "MN"
        If CleanState Like "missi*" Then CleanState = "MS"
        If CleanState Like "misso*" Then CleanState = "MO"
        If CleanState Like "mo*" Then CleanState = "MT"
        If CleanState Like "neb*" Then CleanState = "NE"
        If CleanState Like "nev*" Then CleanState = "NV"
        If CleanState Like "n h*" Then CleanState = "NH"
        If CleanState Like "n j*" Then CleanState = "NJ"
        If CleanState Like "n m*" Then CleanState = "NM"
        If CleanState Like "n y*" Then CleanState = "NY"
        If CleanState Like "n c*" Then CleanState = "NC"
        If CleanState Like "n d*" Then CleanState = "ND"
        If CleanState Like "oh*" Then CleanState = "OH"
        If CleanState Like "ok*" Then CleanState = "OK"
        If CleanState Like "or*" Then CleanState = "OR"
        If CleanState Like "p*" Then CleanState = "PA"
        If CleanState Like "r*" Then CleanState = "RI"
        If CleanState Like "s c*" Then CleanState = "SC"
        If CleanState Like "s d*" Then CleanState = "SD"
        If CleanState Like "ten*" Then CleanState = "TN"
        If CleanState Like "tex*" Then CleanState = "TX"
        If CleanState Like "u*" Then CleanState = "UT"
        If CleanState Like "v*" Then CleanState = "VA"
        If CleanState Like "wa*" Then CleanState = "WA"
        If CleanState Like "w v*" Then CleanState = "WV"
        If CleanState Like "wi*" Then CleanState = "WI"
        If CleanState Like "wy*" Then CleanState = "WY"
        CleanState = Left(UCase(CleanState), 2)

        'AL Alabama      LA Louisiana      OH Ohio
        'AK Alaska  ME Maine  OK Oklahoma
        'AZ Arizona  MD Maryland  OR Oregon
        'AR Arkansas  MA Massachusetts  PA Pennsylvania
        'CA California  MI Michigan  RI Rhode Island
        'CO Colorado  MN Minnesota  SC South Carolina
        'CT Connecticut  MS Mississippi  SD South Dakota
        'DE Delaware  MO  Missouri  TN Tennessee
        'FL Florida  MT Montana  TX Texas
        'GA Georgia  NE Nebraska  UT Utah
        'HI Hawaii  NV Nevada  VT Vermont
        'ID Idaho  NH New Hampshire  VA Virginia
        'IL Illinois  NJ New Jersey  WA  Washington
        'IN Indiana  NM New Mexico  DC Washington, DC
        'IA Iowa  NY New York  WV West Virginia
        'KS Kansas  NC North Carolina  WI Wisconsin
        'KY Kentucky  ND North Dakota  WY Wyoming

    End Function

End Module
