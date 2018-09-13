# Alfred workflow to search the corporate Exchange directory using Outlook

Search for contacts in the Exchange server's global address book using either the contact's name or email address.

Trigger the workflow using `o <contact name>`. When only one contact is found the workflow additionally
returns the organization hierarchy (the contact's manager and any direct reports).

## Optional Configuration

Set `NAME_FILTER` to a regular expression to remove any matching text from the contact's name.
