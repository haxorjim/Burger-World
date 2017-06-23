Attribute VB_Name = "Security"
Function Decrypt(passcomp$, passkey)
a = 0: r = 0: r2 = 1: ccrpty = 0: total = 0: Final$ = ""
cryptkey = passkey / 100

Do Until a = Len(passcomp$)
        a = a + 1
        ccrpty = ccrpty + 1: total = total + 1: intfact = total / 2
        r3 = 0: r = t: r2 = r: t = r2: If r = 1 Then r3 = 1
        If ccrpty >= total Then ccrpty = 0:
        If total >= intfact Then total = 1:
        Decrypt = Decrypt + Chr$(Asc(Mid$(passcomp$, a, 1)) - ccrpty + r3 - cryptkey)
Loop
End Function


Function Encrypt(pass$)
a = 0: r = 0: r2 = 1: ccrpty = 0: total = 0: Final$ = ""
Randomize Timer

Key = Int(Rnd * 500) + 1
Stack.Push Key

cryptkey = Key / 100

Do Until a = Len(pass$)
        a = a + 1
        ccrpty = ccrpty + 1: total = total + 1: intfact = total / 2
        r3 = 0: r = t: r2 = r: t = r2: If r = 1 Then r3 = 1
        If ccrpty >= total Then ccrpty = 0:
        If total >= intfact Then total = 1:
        Encrypt = Encrypt + Chr$(Asc(Mid$(pass$, a, 1)) + ccrpty - r3 + cryptkey)
Loop
        
End Function


