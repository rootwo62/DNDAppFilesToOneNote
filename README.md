# DNDAppFilesToOneNote
Convert DND App Files to OneNote

This app is very fresh and rough.  It takes the files ceryliae created for the Lion's Gate Apps and imports them into OneNote.
ceryliae github is https://github.com/ceryliae.

I was inspired to do this after using the onenote package **crydid** created that contains tons of DND content.  His website is
http://www.cryrid.com/digitaldnd/.  I really liked the layout, but wanted to tailor the templates to my liking, without manually updating every monster.  

With this app you will be able to import files from [Ceryliae's repo](https://github.com/ceryliae) directly into Microsoft OneNote. Most errors occur from template formatting and the app files missing some element.  The debug setting will export page xml files to allow you to review and compare OneNote XML structures.


Requirements:
1. Microsoft OneNote
2. Notepad++ (or any editor)
3. NET 4.0+


Steps to import DNDAppFiles:
- Create a notebook in onenote called 'DND Notebook'.
- Create a section called 'Feats'.
- Add a page called 'Feat Block (Normal)'.
- Make a table that looks like [Feat Block Template](https://imgur.com/a/KuizSFr).
- Update the app config file to reference the notebook, section, templatepage, and file to import.

  - **ONRootPath:** DND Notebook
  - **ONSection:** Feats
  - **ONBlockTemplatePageName:** Feat Block (Normal)
  - **BlockType:** feat
  - **DNDAppFileXML:** Feats.xml
  - **SlowProcess:** false
  - **CopyPageToSourceBookSection:** true
  - **Debug:** false
  

Output pages should look like https://imgur.com/a/dVMrowW.

**ONRootPath** can be split with a / to indicated a section group

 - _e.g. DND Notebook/Monsters_

**BlockTypes:** monster, spell, race, background, feat

**SlowProcess** set to false or anything but true will run all processes at once.

**CopyPageToSourceBookSection** set to true will look for an element called <type> and split it by the commas using the second comma as the "Source" book.  
  
  - _e.g. humanoid (aarakocra), monster manual_ 
  - _if no source is found the **ONSection** will be used._
  - _If the specified section doesn't exsist in the root it will be created_

Currently, the template table cells have to match the examples, the row order doesn't matter as much as the values between the brackets.

Template File Requirements:

- [Race Block Template](https://imgur.com/a/2iQF0f1)
- [Monster Block Template](https://imgur.com/a/czLz9Qp)
- [Spell Block Template](https://imgur.com/a/9rrCI13)
- [Feat Block Template](https://imgur.com/a/KuizSFr)
- [Background Block Template](https://imgur.com/a/7Y2D2Yh)
 
