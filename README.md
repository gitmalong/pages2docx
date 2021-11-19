# pages2docx
AppleScript script to recursively convert all .pages documents within a folder to .docx. Tested on Mac OS Sierra.

## Run script from command line
Convert all documents within ~/Documents/:

`osascript pages2docx.applescript ~/Documents/`

## Notes
- Converted files are stored in the same folder as its originals 
- New filename is the same as the original but the .pages file extension is replaced with .docx
- Existing .docx files with same filename are overwritten

## Todo
- Add parameter for output format
- Support .keynote and .numbers format

## License
MIT
