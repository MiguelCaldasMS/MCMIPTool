# MCMIPTool
Bulk (folder tree) change of MIP classification of files

I had to ditch a much more elegant (and fast) solution because it was a nightmare to authorize within the (understandably) draconian Microsoft IT rules and procedures. I've scaled it back from C++ to PowerShell and from impersonating the identity of the current logged on user to explicitly ask for the current user credentials (OAUTH). Also hardcoded multiple Microsoft Tenant GUIDs instead of the original way of automatically finding the correct identifiers, no matter the tenant logged on to. Ah, well...

A log file with generic messages and two CSV files (one with a list of all the processed files and another one with a list of files that FAILED reclassification) will be left as a side effect :-).

Note that the tool will convert every "Confidential"-classificated document owned by the authenticated user to a "Non-business" classification. You can do this manually, but the tool will do it blindly, WITHOUT ANY CHECKING OR ASKING FOR CONFIRMATION. Use with care. Please.
