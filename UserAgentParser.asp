<%
	  dim browser, browserarr, User_Agent
		User_Agent = Request.ServerVariables("HTTP_USER_AGENT")

		'http://webaim.org/blog/user-agent-string-history/
		'http://www.useragentstring.com/
			
		If instr(UCase(User_Agent),UCase("Chrome")) > 0 Then
			browser = "Google Chrome"
			If instr(UCase(User_Agent),UCase("Chrome/7")) > 0 Then
				browser = "Google Chrome 7"
			Elseif instr(UCase(User_Agent),UCase("Chrome/6")) > 0 Then
				browser = "Google Chrome 6"
			Elseif instr(UCase(User_Agent),UCase("Chrome/5")) > 0 Then
				browser = "Google Chrome 5"
			Elseif instr(UCase(User_Agent),UCase("Chrome/4")) > 0 Then
				browser = "Google Chrome 4"
			Elseif instr(UCase(User_Agent),UCase("Chrome/3")) > 0 Then
				browser = "Google Chrome 3"
			Elseif instr(UCase(User_Agent),UCase("Chrome/2")) > 0 Then
				browser = "Google Chrome 2"
			End If
			
			browserarr = Split(    Replace(   right(User_Agent, (len(User_Agent)+1-instr(UCase(User_Agent),UCase("Chrome"))))   , "Chrome/","")      )
			browser = "Google Chrome " & browserarr(0)
			
		Elseif instr(UCase(User_Agent),UCase("Safari")) > 0 Then
			browser = "Safari"
			If instr(UCase(User_Agent),UCase("Version/5")) > 0 Then
				browser = "Safari 5"
			Elseif instr(UCase(User_Agent),UCase("Version/4")) > 0 Then
				browser = "Safari 4"
			Elseif instr(UCase(User_Agent),UCase("Version/3")) > 0 Then
				browser = "Safari 3"
			End If
			
			browserarr = Split(    Replace(   right(User_Agent, (len(User_Agent)+1-instr(UCase(User_Agent),UCase("version"))))   , "Version/","")      )
			browser = "Safari " & browserarr(0)
				
		Elseif instr(UCase(User_Agent),UCase("Opera")) > 0 Then
			browser = "Opera"
			If instr(UCase(User_Agent),UCase("Opera/11")) > 0 Then
				browser = "Opera 11"
			Elseif instr(UCase(User_Agent),UCase("Opera/10")) > 0 Then
				browser = "Opera 10"
			Elseif instr(UCase(User_Agent),UCase("Opera/9.5")) > 0 Then
				browser = "Opera 9.5"
			Elseif instr(UCase(User_Agent),UCase("Opera/9")) > 0 Then
				browser = "Opera 9"
			End If
			
			browserarr = Split(    Replace(   right(User_Agent, (len(User_Agent)+1-instr(UCase(User_Agent),UCase("Opera"))))   , "Opera/","")      )
			browser = "Opera " & browserarr(0)
			
		ElseIf instr(UCase(User_Agent),"MSIE") > 0 Then
			browser = "Microsoft Internet Explorer"
			If instr(UCase(User_Agent),"MSIE 9") > 0 Then
				browser = "Microsoft Internet Explorer 9"
			Elseif instr(UCase(User_Agent),"MSIE 8") > 0 Then
				browser = "Microsoft Internet Explorer 8"
			Elseif instr(UCase(User_Agent),"MSIE 7") > 0 Then
				browser = "Microsoft Internet Explorer 7"
			Elseif instr(UCase(User_Agent),"MSIE 6") > 0 Then
				browser = "Microsoft Internet Explorer 6"
			End If	
			
			browserarr = Split(    Replace(   right(User_Agent, (len(User_Agent)+1-instr(UCase(User_Agent),UCase("MSIE"))))   , "MSIE ","")      )
			browser = "Microsoft Internet Explorer " &  Replace(browserarr(0), ";", "")
			
		Elseif instr(UCase(User_Agent),UCase("Firefox")) > 0 Then
			browser = "Mozilla Firefox"
			If instr(UCase(User_Agent),UCase("Firefox/4")) > 0 Then
				browser = "Mozilla Firefox 4"
			Elseif instr(UCase(User_Agent),UCase("Firefox/3")) > 0 Then
				browser = "Mozilla Firefox 3"
			Elseif instr(UCase(User_Agent),UCase("Firefox/2")) > 0 Then
				browser = "Mozilla Firefox 2"
			End If
			
			browserarr = Split(    Replace(   right(User_Agent, (len(User_Agent)+1-instr(UCase(User_Agent),UCase("Firefox"))))   , "Firefox/","")      )
			browser = "Mozilla Firefox " & Replace(browserarr(0), "Firefox", "")
			
			
		Else
			browser = User_Agent
		End If






		Dim OS
		If instr(UCase(User_Agent),UCase("Windows")) > 0 Then
			OS = "Windows"
			If instr(UCase(User_Agent),UCase("Windows NT 6.1")) > 0 Then
				OS = "Windows 7"
			Elseif instr(UCase(User_Agent),UCase("Windows NT 6.0")) > 0 Then
				OS = "Windows Vista"
			Elseif (instr(UCase(User_Agent),UCase("Windows NT 5.1")) > 0) or (instr(UCase(User_Agent),UCase("Windows NT 5.2")) > 0)Then
				OS = "Windows XP"
			End If
		Elseif instr(UCase(User_Agent),UCase("Macintosh")) > 0 Then
			OS = "Macintosh"
		Elseif instr(UCase(User_Agent),UCase("Linux")) > 0 Then
			OS = "Linux"
			If instr(UCase(User_Agent),UCase("Ubuntu")) > 0 Then
				OS = "Linux Ubuntu"
			End If
		Elseif instr(UCase(User_Agent),UCase("SunOS")) > 0 Then
			OS = "Solaris - SunOS"
		Elseif instr(UCase(User_Agent),UCase("FreeBSD")) > 0 Then
			OS = "FreeBSD"
		Else
			OS = User_Agent
		End If



		Dim mobile_device
		If (instr(UCase(User_Agent),UCase("Android")) > 0) OR (instr(UCase(User_Agent),UCase("Windows Phone")) > 0) OR (instr(UCase(User_Agent),UCase("Phone")) > 0) OR (instr(UCase(User_Agent),UCase("iPhone")) > 0) OR (instr(UCase(User_Agent),UCase("iPad")) > 0) OR (instr(UCase(User_Agent),UCase("iPod")) > 0) OR (instr(UCase(User_Agent),UCase("Kindle")) > 0) OR (instr(UCase(User_Agent),UCase("Blackberry")) > 0) Then
			moblie_device = true
			OS = User_Agent
			browser = User_Agent
		End If

		'http://www.useragentstring.com/pages/api.php
		'http://www.useragentstring.com/?uas=User_Agent&getText=all
		'http://www.useragentstring.com/?uas=User_Agent&getJSON=all
		'http://www.zytrax.com/tech/web/mobile_ids.html
		
%>
