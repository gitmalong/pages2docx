-- This script recursively traverses a directory and converts all .pages documents to .docx documents
-- Tested on Mac OS Sierra
-- Command line usage: osascript pages2docx.applescript ~/Documents/
on run inputFolder
	
	tell application "Finder" to set theFiles to every file in the entire contents of folder inputFolder
	repeat with theFile in theFiles
		set fExt to name extension of theFile as text
		set fName to name of theFile as text
		set theFilesFolder to folder of theFile as text
		
		if fExt is "pages" then
			
			tell application "Finder"
				set theFilesFolder to (folder of theFile) as text
			end tell
			
			tell application "Pages"
				set theDoc to open (theFilesFolder & fName)
				set theDocName to name of theDoc
				set thePDFPath to (theFilesFolder & theDocName & ".docx") as text
				close access (open for access thePDFPath)
				export theDoc to file thePDFPath as Microsoft Word
				close theDoc
				
			end tell
			
			tell application "Finder"
				move file thePDFPath to folder theFilesFolder with replacing
			end tell
		end if
		
	end repeat
	
end run