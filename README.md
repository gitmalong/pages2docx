# pages2docx
AppleScript script to recursively convert all .pages documents within a folder to .docx. Tested on Mac OS Sierra.

## Usage with Automator
- Create new application
- Drag a "Run AppleScript" action into your view 
- Replace the existing code with the content of pages2docx.applescript
- Drag a folder that contains the documents into your view
- Run script

## Notes
- Converted files are stored in the same folder as its originals 
- New filename is the same as the original but the .pages file extension is replaced with .docx
- Existing .docx files with same filename are overwritten

## License
MIT
