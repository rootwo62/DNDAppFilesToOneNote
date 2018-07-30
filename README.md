# DNDAppFilesToOneNote
Convert DND App Files to OneNote

This app is very fresh and rough.  It takes the files ceryliae created for the Lion's Gate Apps and imports them into OneNote.
ceryliae github is https://github.com/ceryliae.

I was inspired to do this after using the onenote package crydid created that contains tons of DND content.  His website is
http://www.cryrid.com/digitaldnd/.


Requirements:
1. Microsoft OneNote
2. Notepad++ (or any editor)
3. NET 4.0+


Steps to import monsters:
- Create a notebook in onenote called 'Development'.
- Create a section called 'temporary'.
- Add a page called 'Monster Block (Normal).
- Make a table that looks like https://imgur.com/a/C8nPxFk.
- Update the app config file to reference the notebook, section, page, and file to import.
  - **Notebook:** Development
  - **Section:** temporary
  - **MonsterBlockTemplate:** Monster Block (Normal)
  - **BlockType:** monster
- Output pages should look like https://imgur.com/a/dVMrowW.

_supported BlockTypes: monster, spell, race_
