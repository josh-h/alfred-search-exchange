on run argv
    set {text item delimiters, TID} to {" ", text item delimiters}
    set {text item delimiters, query} to {TID, argv as text}
    
    tell application "Microsoft Outlook"
		set exchange to default account
		tell shared contacts panel
			set search string to query
			
			repeat while (get searching)
				delay 0.05
			end repeat
			
			set response to ""
			if (count of contacts) > 0 then
				repeat with i from 1 to (count of contacts)
					
					set response to response & "contact:\"" & (display name of contact i) & "\""
					
					set response to response & ", title:\"" & job title of contact i & "\""

					set emails to email addresses of contact i
					repeat with email in emails
						set response to response & ", address:\"" & address of email & "\", "
					end repeat

				end repeat
			end if
			return response
		end tell
    end tell
end run
