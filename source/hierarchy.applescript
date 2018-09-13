on run argv
    tell application "Microsoft Outlook"
        set exchange to default account
        tell shared contacts panel
            set search string to argv

            repeat while (get searching)
                delay 0.1
            end repeat
            
            set theContact to first item of contacts
            set org to query organization information exchange for theContact
            set response to ""
            set theManagers to (get manager of org)
            repeat with theManager in theManagers
                set response to response & "manager:\"" & (get the name of theManager)  & "\""
                set response to response & ", title:\"" & (get the title of theManager) & "\" "
            end repeat

            set reports to (get direct reports of org)
            repeat with direct in reports
                set response to response & "direct_report:\"" & (get the name of the direct) & "\""
                set response to response & ", title:\"" & (get the title of the direct) & "\" "
            end repeat

            return response
        end tell

    end tell
end run
