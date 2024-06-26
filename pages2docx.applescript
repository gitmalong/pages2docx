on run inputFolder
	
	-- Define a list to hold the names of files that could not be processed
	set unprocessedFiles to {}
	-- Define a list to hold the names of files that were processed successfully
	set processedFiles to {}
	
	tell application "Finder" to set theFiles to every file in the entire contents of folder inputFolder
	
	repeat with theFile in theFiles
		
		set fExt to name extension of theFile as text
		set fName to name of theFile as text
		set fDir to folder of theFile as text
		
		if fExt is "pages" then
			using terms from application "Pages"
				convert(fDir, fName, ".docx", unprocessedFiles, processedFiles)
				
			end using terms from
		end if
		
	end repeat
	
	-- Print the list of unprocessed files at the end of the script
	log "Unprocessed files:\n" & my join(unprocessedFiles, "\n")
	-- Print the list of successfully processed files at the end of the script
	log "Successfully processed files:\n" & my join(processedFiles, "\n")
	
end run

on convert(dirName, fileName, exportExtension, unprocessedFiles, processedFiles)
	
	tell application "Pages"
		set fullPath to (dirName & fileName)
		set posixFullPath to POSIX path of fullPath
		set exportFileName to (dirName & fileName & exportExtension) as text

		try
			set doc to open fullPath
			if doc is missing value then error "Document is missing"
			set docName to name of doc

			close access (open for access exportFileName)

			tell application "Pages"
				try
					export doc to file exportFileName as Microsoft Word
					-- If the export succeeds, add the full POSIX path to the list of processed files
					set end of processedFiles to posixFullPath

					tell application "Finder"
						move file exportFileName to folder dirName with replacing
					end tell
				on error
					-- If an error occurs, add the full POSIX path to the list of unprocessed files
					set end of unprocessedFiles to posixFullPath
				end try
			end tell
			
			close doc
		on error
			-- If an error occurs during opening, add the full POSIX path to the list of unprocessed files
			set end of unprocessedFiles to posixFullPath
		end try
	end tell
	

	
end convert

on join(theList, theDelimiter)
	set theString to ""
	set theOldDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to theDelimiter
	set theString to theList as string
	set AppleScript's text item delimiters to theOldDelimiters
	return theString
end join