Function getPassword(digits As Integer) As String
	Dim i As Integer, s As String, sCharList As String
	Dim bUpper As Boolean, bLower As Boolean, bNumber As Boolean
	
	Randomize '乱数系列の再設定
	
	' 小文字のエル、オー、大文字のアイ、オー、数字のゼロ、イチは紛らわしいので候補から除く
	sCharList = "abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789"
	
	Do
		getPassword = "": bUpper = False: bLower = False: bNumber = False
		For i = 1 To digits
			s = Mid(sCharList, Int(1 + Rnd * Len(sCharList)), 1)
			If s >= "A" And s <= "Z" Then bUpper = True
			If s >= "a" And s <= "z" Then bLower = True
			If s >= "0" And s <= "9" Then bNumber = True
			getPassword = getPassword & s
		Next
	Loop Until bUpper And bLower And bNumber
End Function
