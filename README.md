This is a VBScript function for measuring password entropy based on the algorithm used by passwordmeter.com (released under a GPL license)

The function will generate a score (0% - 100%) based on the strength of a user's password.

Examples:

	qwerty (0%)
	password (3%)
	password123 (26%) 
	p455w0rd (37%)
	PassWord123 (49%)
	PaSSword658 (51%)
	P@$$w0rd (58%)
	p@4$W0rD743! (95%)

To use the function as a means of validating signup requests, simply reject any requests where the user's password fails to meet a minimum strength percentage (E.g: 40%)
