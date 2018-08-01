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


Steps to import DNDAppFiles:
- Create a notebook in onenote called 'DND Notebook'.
- Create a section called 'Feats'.
- Add a page called 'Feat Block (Normal).
- Make a table that looks like [Feat Block Template](https://imgur.com/a/KuizSFr).
- Update the app config file to reference the notebook, section, templatepage, and file to import.
  - **Notebook:** DND Notebook
  - **Section:** Feats
  - **BlockTemplatePageName:** Feat Block (Normal)
  - **BlockType:** feat
  - **DNDAppFileXML:** Feats.xml
  - **NamedList:** Actor, Alert
- Output pages should look like https://imgur.com/a/dVMrowW.

_supported BlockTypes: monster, spell, race, background, feat_
_leave __NamedList__ setting blank to get all items in xml document_

Currently, the template table cells have to match the examples, the row order doesn't matter as much as the values between the brackets.

Template File Requirements:
- [Race Block Template](https://imgur.com/a/2iQF0f1)
- [Monster Block Template](https://imgur.com/a/czLz9Qp)
- [Spell Block Template](https://imgur.com/a/9rrCI13)
- [Feat Block Template](https://imgur.com/a/KuizSFr)
- [Background Block Template](https://imgur.com/a/7Y2D2Yh)
