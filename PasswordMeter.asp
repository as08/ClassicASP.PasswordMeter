<%

	Const min_pw_len = 6
	Const max_pw_len = 100

	function password_strength(ByVal pwd)		
		
		if isEmpty(pwd) OR isNull(pwd) then
			
			password_strength = 0
			exit function
			
		end if
		
		'------------------------------------------'
		' Remove any white space, but leave spaces '
		'------------------------------------------'
					
		Dim RegEx : Set RegEx = New RegExp
		
			RegEx.Pattern = "[^\S ]+"
			RegEx.Multiline = True
			RegEx.Global = True
			pwd = RegEx.Replace(pwd,"")
		
		Set RegEx = nothing			
		
		'------------------------------------------------------------------------------------------------'
		' Reject the password if it doesn't meet the min/max length requirements and return a score of 0 '                     '
		'------------------------------------------------------------------------------------------------'
			
		if NOT (len(pwd) >= min_pw_len AND len(pwd) <= max_pw_len) then
			
			password_strength = 0
			exit function
		
		end if
			
		'--------------------------------------------'
		' Create an array of the password characters '
		'--------------------------------------------'

		Dim a, b, c, arrPwd()

		ReDim arrPwd(len(pwd) - 1)

		for a = 1 to len(pwd)
			arrPwd(a - 1) = mid(pwd,a,1)
		next

		'-----------------------------------------------------------------------------------'
		' Get the password length after the white space removal for a more accurate reading '
		'-----------------------------------------------------------------------------------'

		Dim length : length = uBound(arrPwd) + 1
		Dim score : score = int(length * 4)

		'---------------------------------------------------------------------------------------'
		' score will be converted to a percentage (0-100), but to calculate a percentage we     '
		' first need to set a target. The target should be the score value of what you consider '
		' to be a "very strong" password. This will be the benchmark for all password scores.   '
		' The higher the value the lower the percentage will be for weaker passwords.           '
		'---------------------------------------------------------------------------------------'

		Dim scoreTarget : scoreTarget = 150

		'----------------------------------------------------------------------------------'
		' Set a recommended minimum password length and penalize passwords that fall short '                           '
		'----------------------------------------------------------------------------------'

		Dim recMinPwdLen : recMinPwdLen = 8

		'---------------------------'
		' Set the default variables '
		'---------------------------'

		Dim alphaUC : alphaUC = 0
		Dim alphaLC : alphaLC = 0
		Dim numbers : numbers = 0
		Dim symbols : symbols = 0
		Dim midChar : midChar = 0				
		Dim unqChar : unqChar = 0
		Dim repChar : repChar = 0
		Dim repInc : repInc = 0
		Dim minReq : minReq = 0
		Dim requirements : requirements = 0

		Dim consecAlphaUC : consecAlphaUC = 0
		Dim consecAlphaLC : consecAlphaLC = 0
		Dim consecNumber : consecNumber = 0
		Dim consecSymbol : consecSymbol = 0

		Dim seqAlpha : seqAlpha = 0
		Dim seqNumber : seqNumber = 0
		Dim seqSymbol : seqSymbol = 0

		Dim multMidChar : multMidChar = 2
		Dim multRequirements : multRequirements = 2
		Dim multConsecAlphaUC : multConsecAlphaUC = 2
		Dim multConsecAlphaLC : multConsecAlphaLC = 2
		Dim multConsecNumber : multConsecNumber = 2
		Dim multConsecSymbol : multConsecSymbol = 1

		Dim multSeqAlpha : multSeqAlpha = 3
		Dim multSeqNumber : multSeqNumber = 3
		Dim multSeqSymbol : multSeqSymbol = 3

		Dim multNumber : multNumber = 4
		Dim multSymbol : multSymbol = 6

		Dim tmpAlphaUC : tmpAlphaUC = ""
		Dim tmpAlphaLC : tmpAlphaLC = ""
		Dim tmpNumber : tmpNumber = ""
		Dim tmpSymbol : tmpSymbol = ""

		Dim alphaPtrn : alphaPtrn = "abcdefghijklmnopqrstuvwxyzqwertyuiopasdfghjklzxcvbnm"
		Dim numericPtrn : numericPtrn = "01234567890"
		Dim symbolPtrn : symbolPtrn = "¬!""£$%^&*():@~;'#,./<>?"

		Dim charExists : charExists = false

		'---------------------------------------------------'
		' Set the character type regular expression objects '
		'---------------------------------------------------'

		Dim RegExUC, RegExLC, RegExNumber, RegExSymbol

		Set RegExUC = New RegExp

			RegExUC.Pattern = "[A-Z]" ' Look for uppercase letters
			RegExUC.Global = True

		Set RegExLC = New RegExp

			RegExLC.Pattern = "[a-z]" ' Look for lowercase letters
			RegExLC.Global = True

		Set RegExNumber = New RegExp

			RegExNumber.Pattern = "[0-9]" ' Look for lowercase numbers
			RegExNumber.Global = True

		Set RegExSymbol = New RegExp

			RegExSymbol.Pattern = "[^a-zA-Z0-9_]" ' Look for symbols
			RegExSymbol.Global = True

		'----------------------------------------------------------------------------------------'
		' Loop through password to check for Symbol, Numeric, Lowercase and Uppercase characters '
		'----------------------------------------------------------------------------------------'

		for a = 0 to uBound(arrPwd)

			'-------------------------------------------------------------------------------------------'
			' Check each character to determine its type.                                               '
			'                                                                                           '
			' Keep a character type count as well as a count of consecutive character types.            '
			' For example: "password123" has 8 consecutive lowercase letters and 3 consecutive numbers. '
			'                                                                                           '
			' Also check for numbers and symbols that appear in the middle of passwords. This is often  '
			' an indicator of good password entropy.                                                    '
			'-------------------------------------------------------------------------------------------'

			'----------------------------'
			' Is it an uppercase letter? '
			'----------------------------'

			if RegExUC.Test(arrPwd(a)) then

				if NOT tmpAlphaUC = "" then if (tmpAlphaUC + 1) = a then consecAlphaUC = consecAlphaUC + 1

				tmpAlphaUC = a
				alphaUC = alphaUC + 1

			'---------------------------'
			' Is it a lowercase letter? '
			'---------------------------'

			elseif RegExLC.Test(arrPwd(a)) then

				if NOT tmpAlphaLC = "" then if (tmpAlphaLC + 1) = a then consecAlphaLC = consecAlphaLC + 1

				tmpAlphaLC = a
				alphaLC = alphaLC + 1

			'-----------------'
			' Is it a number? '
			'-----------------'

			elseif RegExNumber.Test(arrPwd(a)) then

				if a > 0 AND a < uBound(arrPwd) then midChar = midChar + 1

				if NOT tmpNumber = "" then if (tmpNumber + 1) = a then consecNumber = consecNumber + 1

				tmpNumber = a
				numbers = numbers + 1

			'-----------------'
			' Is it a symbol? '
			'-----------------'

			elseif RegExSymbol.Test(arrPwd(a)) then

				if a > 0 AND a < uBound(arrPwd) then midChar = midChar + 1

				if NOT tmpSymbol = "" then if (tmpSymbol + 1) = a then consecSymbol = consecSymbol + 1

				tmpSymbol = a
				symbols = symbols + 1

			end if

			'------------------------------------------------------------------------------------------'
			' After analysing the character type, create a second loop to check for repeat characters. '
			' Calculate an increment deduction based on the proximity to identical characters. The     '
			' deduction is incremented each time a new match is discovered. The deduction amount is    '
			' based on the total password length divided by the difference in distance between the     '
			' currently selected match.                                                                '
			'------------------------------------------------------------------------------------------'

			charExists = false

			for b = 0 to uBound(arrPwd)
				if arrPwd(a) = arrPwd(b) AND NOT a = b then
					charExists = true
					repInc = repInc + abs(length / (b - a))
				end if
			next

			'------------------------------------------------------------'
			' Keep count of the number of repeated and unique characters '
			'------------------------------------------------------------'

			if charExists then

				repChar = repChar + 1
				unqChar = length - repChar

				'-------------------------------------------------------------------------------------------'
				' Divide the increment deduction for repeated characters aginst the unique character count  '
				' and round up. Check the the unique count is greater than 0 to avoid division by 0 errors. '
				' If the unique count is 0 and the increment deduction is a decimal then round up.          '
				'-------------------------------------------------------------------------------------------'

				if unqChar > 0 then

					if repInc MOD unqChar = 0 Then
						repInc = repInc/unqChar
					else
						repInc = int(repInc / unqChar) + 1
					end if

				elseif varType(repInc) = 5 then

					repInc = int(repInc) + 1 ' Round up

				end if
			end if

		next

		'-----------------------------------------------------'
		' Clear the character type regular expression objects '
		'-----------------------------------------------------'

		Set RegExUC = nothing
		set RegExLC = nothing
		Set RegExNumber = nothing
		Set RegExSymbol = nothing

		'-------------------------------------------------------------------------------------'	
		' Look for sequential patterns, both forward and reverse in groups of 3. For symbols, '
		' a pattern is defined as 3 or more symbols where the symbol keys are situated next   '
		' to each other on a qwerty keyboard                                                  '
		'-------------------------------------------------------------------------------------'

		Dim fwd, rev

		'------------------------------------------------------------------'
		' Check for sequential alpha string patterns (forward and reverse) '
		'------------------------------------------------------------------'

		for c = 1 to len(alphaPtrn) - 3
			fwd = mid(alphaPtrn,c,3) : rev = StrReverse(fwd)
			if inStr(lCase(pwd),fwd) > 0 OR inStr(lCase(pwd),rev) > 0 then seqAlpha = seqAlpha + 1
		next

		'--------------------------------------------------------------------'
		' Check for sequential numeric string patterns (forward and reverse) '
		'--------------------------------------------------------------------'

		for c = 1 to len(numericPtrn) - 3
			fwd = mid(numericPtrn,c,3) : rev = StrReverse(fwd)
			if inStr(lCase(pwd),fwd) > 0 OR inStr(lCase(pwd),rev) > 0 then seqNumber = seqNumber + 1
		next

		'-------------------------------------------------------------------'
		' Check for sequential symbol string patterns (forward and reverse) '
		'-------------------------------------------------------------------'

		for c = 1 to len(symbolPtrn) - 3
			fwd = mid(symbolPtrn,c,3) : rev = StrReverse(fwd)
			if inStr(pwd,fwd) > 0 OR inStr(pwd,rev) > 0 then seqSymbol = seqSymbol + 1
		next

		'--------------------------------------------------------------'
		' Modify overall score value based on usage vs requirements.   '
		' Requirement points are assigned for the following:           '
		'                                                              '
		' - Passwords that contain uppercase letters                   '
		' - Passwords that contain lowercase letters                   '
		' - Passwords that contain numbers                             '
		' - Passwords that contain symbols                             '
		'                                                              '
		' Requirement points are doubled and added to the score, but   '
		' only if the minimum number of requirements are met.          '
		'                                                              '
		' For passwords that are 8+ characters the minimum requirement '
		' is 2. For passwords less than 8 characters the minimum is 4. '
		'--------------------------------------------------------------'

		' General point assignment

		if alphaUC > 0 AND alphaUC < length then score = int(score + ((length - alphaUC) * 2)) : _
		requirements = requirements + 1	

		if alphaLC > 0 AND alphaLC < length then score = int(score + ((length - alphaLC) * 2)) : _
		requirements = requirements + 1

		if numbers > 0 AND numbers < length then score = int(score + (numbers * multNumber)) : _
		requirements = requirements + 1

		if symbols > 0 then score = int(score + (symbols * multSymbol)) : _
		requirements = requirements + 1

		if midChar > 0 then score = int(score + (midChar * multMidChar))

		'-------------------------------------'
		' Point deductions for poor practices '
		'-------------------------------------'

		' Only Letters
		if (alphaLC > 0 OR alphaUC > 0) AND symbols = 0 AND numbers = 0 then score = int(score - length)

		' Only Numbers
		if alphaLC = 0 AND alphaUC = 0 AND symbols = 0 AND numbers > 0 then score = int(score - length)

		' Same character exists more than once	
		if repChar > 0 then score = int(score - repInc)

		' Consecutive Uppercase Letters exist
		if consecAlphaUC > 0 then score = int(score - (consecAlphaUC * multConsecAlphaUC))

		' Consecutive Lowercase Letters exist	
		if consecAlphaLC > 0 then score = int(score - (consecAlphaLC * multConsecAlphaLC))

		' Consecutive Numbers exist
		if consecNumber > 0 then score = int(score - (consecNumber * multConsecNumber))

		' Consecutive Sumbols exist
		if consecSymbol > 0 then score = int(score - (consecSymbol * multConsecSymbol))

		' Sequential alpha strings exist (3 characters or more)
		if seqAlpha > 0 then score = int(score - (seqAlpha * multSeqAlpha))

		' Sequential numeric strings exist (3 characters or more)
		if seqNumber > 0 then score = int(score - (seqNumber * multSeqNumber))

		' Sequential symbol strings exist (3 characters or more)
		if seqSymbol > 0 then score = int(score - (seqSymbol * multSeqSymbol))

		'--------------------------------------------------------'
		' Increase the score if the minimum requirements are met '
		'--------------------------------------------------------'

		if length >= recMinPwdLen then minReq = 2 else minReq = 4

		if requirements > minReq then score = int(score + (requirements * multRequirements))

		'---------------------------------------'
		' Return a strength score between 0-100 '
		'---------------------------------------'

		if score > scoreTarget then score = scoreTarget
		if score < 0 then score = 0

		score = int(score / scoreTarget * 100)

		'---------------------------------------------'
		' SCORE TRANSLATIONS                          '
		'---------------------------------------------'
		' score >= 0  AND score <  20   = Very Weak   '
		' score >= 20 AND score <  40   = Weak        '
		' score >= 40 AND score <  60   = Good        '
		' score >= 60 AND score <  80   = Strong      '
		' score >= 80 AND score <= 100  = Very Strong '
		'---------------------------------------------'

		' Average execution time : 0.0085s

		password_strength = score
				
	end function

%>
