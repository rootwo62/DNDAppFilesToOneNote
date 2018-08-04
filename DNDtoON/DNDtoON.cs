using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Threading;
using OneNote = Microsoft.Office.Interop.OneNote;


namespace DNDtoON
{
    class Application
    {

        OneNote.Application OneNoteApp = new OneNote.Application();

        string notebookID, sectionID, pageID;
        string notebooksxmlfile, sectionsxmlfile, pagesxmlfile;
        string sectionname, notebookName;

        const string font = @"<span style = 'font-size:9pt; font-family:cambria'>";
        const string boldfont = @"<span style = 'font-weight:bold; font-size:9pt; font-family:cambria'>";
        const string boldname = @"<span style = 'font-weight:bold; font-size:12pt; font-family:cambria'>";
        const string tableheaderblack = @"<span style = 'color:#FFFFFF; font-weight:bold; font-size:9pt; font-family:cambria;'>";
        const string tablefont = @"<span style = 'color:#000000; font-size:8pt; font-family:cambria;'>";
        const string italicfont = @"<span style='font-style:italic; font-size:9pt; font-family:cambria'>";
        const string bolditalicfont = @"<span style='font-style:italic; font-weight:bold; font-size:9pt; font-family:cambria'>";
        const string endspan = @"</span>";

        static void Main(string[] args)
        {

            Application Application = new Application();
            Application.Run();

            //Application.Run();

            Console.WriteLine("press any key to close...");
            Console.ReadKey(true);
        }

        private void Run()
        {
            string Compendium = Properties.Settings.Default.DNDAppFileXML;

            XDocument xdocCompendium = XDocument.Load(Compendium);

            string blocktype = Properties.Settings.Default.BlockType;
            string templatexmlfile = blocktype + "templatefile.xml";
            GetOneNoteTableXML(Properties.Settings.Default.BlockTemplatePageName, templatexmlfile);

            Console.WriteLine("NOTEBOOK: {1}{0}SECTION: {2}{0}COMPENDIUM: {3}{0}Press any key to continue...{0}", Environment.NewLine, notebookName, sectionname, templatexmlfile);
            Console.ReadKey(true);

            foreach (XElement element in xdocCompendium.Descendants(blocktype))
            {
                string pageName = element.DescendantsAndSelf("name").First().Value;
                Console.WriteLine("adding {0} to section {1} in notebook {2}", pageName, sectionname, notebookName);
                CopyPageTableTemplate(pageName, templatexmlfile, out string pageID);

                string node = element.ToString();

                if (blocktype == "monster")
                    FillMonsterStatsTable(node, pageID);
                if (blocktype == "spell")
                    FillSpellTable(node, pageID);
                if (blocktype == "race")
                    FillRaceTable(node, pageID);
                if (blocktype == "feat")
                    FillFeatTable(node, pageID);
                if (blocktype == "background")
                    FillBackgroundTable(node, pageID);
                if (blocktype == "class")
                    FillClassTable(node, pageID);

                if (Properties.Settings.Default.SlowProcess)
                    Console.ReadKey(true);
            }
        }

        private void GetONHeierarchy(string ID, OneNote.HierarchyScope OneNoteHierarchyScope, string OutputFile)
        {
            OneNoteApp.GetHierarchy(ID, OneNoteHierarchyScope, out string pbstrHierarchyXMLOut);
            XDocument.Parse(pbstrHierarchyXMLOut).Save(OutputFile);
        }

        private void CreateONPage(string PageName, string ONHierarchyID, out string NewPageID, out string NewPageXML)
        {
            XNamespace URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            OneNoteApp.CreateNewPage(ONHierarchyID, out NewPageID);
            OneNoteApp.GetPageContent(NewPageID, out NewPageXML);

            XDocument xdocNewPage = XDocument.Parse(NewPageXML);
            xdocNewPage.Root.Element(URI + "Title").Descendants(URI + "T").First().Value = PageName;

            OneNoteApp.UpdatePageContent(xdocNewPage.ToString());
            OneNoteApp.SyncHierarchy(sectionID);
        }

        // Load Variables For Specified Block

        private void GetOneNoteTableXML(string templatePageName, string outXMLFile)
        {
            sectionname = Properties.Settings.Default.Section;
            notebookName = Properties.Settings.Default.Notebook;

            notebooksxmlfile = "notebooks.xml";
            sectionsxmlfile = "sections.xml";
            pagesxmlfile = "pages.xml";

            GetONHeierarchy("", OneNote.HierarchyScope.hsNotebooks, notebooksxmlfile);
            notebookID = ID(notebooksxmlfile, notebookName);

            GetONHeierarchy(notebookID, OneNote.HierarchyScope.hsSections, sectionsxmlfile);
            sectionID = ID(sectionsxmlfile, sectionname);

            GetONHeierarchy(sectionID, OneNote.HierarchyScope.hsPages, pagesxmlfile);
            pageID = ID(pagesxmlfile, templatePageName);

            OneNoteApp.GetHierarchy(pageID, OneNote.HierarchyScope.hsChildren, out string tableTemplateXML);

            XDocument.Parse(tableTemplateXML).Save(outXMLFile);
        }

        // Copy Templates

        private void CopyPageTableTemplate(string outPageName, string templatePageXML, out string pageID)
        {
            XNamespace URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            OneNoteApp.CreateNewPage(sectionID, out pageID);
            OneNoteApp.GetPageContent(pageID, out string NewPageXML);
            XDocument xdocTemplate = XDocument.Load(templatePageXML);
            XElement xelTemplateOutline = xdocTemplate.Root.Element(URI + "Outline");
            xelTemplateOutline.Attribute("objectID").Remove();
            XDocument docNewPage = XDocument.Parse(NewPageXML);
            docNewPage.Root.Add(xelTemplateOutline);
            docNewPage.Root.Descendants(URI + "Title").First().Descendants(URI + "T").First().Value = outPageName;

            try
            {
                OneNoteApp.UpdatePageContent(docNewPage.ToString());
                OneNoteApp.SyncHierarchy(sectionID);
            }
            catch (Exception ex)
            {
                Console.WriteLine("failed to copy template for {0}", outPageName);
                Console.WriteLine(ex.Message.Trim());
                throw;
            }

        }

        // Fill Tables

        private void FillMonsterStatsTable(string InputXML, string MonsterPageID)
        {

            XmlDocument Monster = new XmlDocument();
            Monster.LoadXml(InputXML);

            OneNoteApp.GetHierarchy(MonsterPageID, OneNote.HierarchyScope.hsChildren, out string CurrentMonsterPageXml);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument MonsterPageXMLDoc = new XmlDocument();
            MonsterPageXMLDoc.LoadXml(CurrentMonsterPageXml);
            XmlNamespaceManager XNSMGR = XMLMGR(MonsterPageXMLDoc);
            XmlNode Table = MonsterPageXMLDoc.SelectSingleNode("//one:Table", XNSMGR);

            // single child nodes
            string name = Monster.SelectSingleNode("//name")?.InnerText;
            string size = Monster.SelectSingleNode("//size")?.InnerText;
            string type = Monster.SelectSingleNode("//type")?.InnerText;
            string alignment = Monster.SelectSingleNode("//alignment")?.InnerText;
            string ac = Monster.SelectSingleNode("//ac")?.InnerText;
            string hp = Monster.SelectSingleNode("//hp")?.InnerText;
            string speed = Monster.SelectSingleNode("//speed")?.InnerText;
            string strength = Monster.SelectSingleNode("//str")?.InnerText;
            string dexterity = Monster.SelectSingleNode("//dex")?.InnerText;
            string constitution = Monster.SelectSingleNode("//con")?.InnerText;
            string intelligence = Monster.SelectSingleNode("//int")?.InnerText;
            string wisdom = Monster.SelectSingleNode("//wis")?.InnerText;
            string charisma = Monster.SelectSingleNode("//cha")?.InnerText;
            string save = Monster.SelectSingleNode("//save")?.InnerText;
            string skill = Monster.SelectSingleNode("//skill")?.InnerText;
            string resist = Monster.SelectSingleNode("//resist")?.InnerText;
            string immune = Monster.SelectSingleNode("//immune")?.InnerText;
            string conditionimmune = Monster.SelectSingleNode("//conditionImmune")?.InnerText;
            string senses = Monster.SelectSingleNode("//senses")?.InnerText;
            string passive = Monster.SelectSingleNode("//passive")?.InnerText;
            string languages = Monster.SelectSingleNode("//languages")?.InnerText;
            string cr = ChallengeRating(Monster.SelectSingleNode("//cr")?.InnerText);
            string spells = Monster.SelectSingleNode("//spells")?.InnerText;
            string legendary = Monster.SelectSingleNode("//legendary")?.InnerText;
            string traits = Monster.SelectSingleNode("//trait")?.InnerText;
            string actions = Monster.SelectSingleNode("//action")?.InnerText;
            string enviornment = "any";

            size = MonsterSize(size);

            // name
            if (name != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = boldname + name + endspan;
            // size
            if (size != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[size]')]", XNSMGR).InnerText = italicfont + size + ", " + type + ", " + alignment + endspan;
            // armor class
            if (ac != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[ac]')]", XNSMGR).InnerText = boldfont + @"Armor Class: " + endspan + ac;
            // hit points
            if (hp != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[hp]')]", XNSMGR).InnerText = boldfont + @"Hit Points: " + endspan + hp;
            // speed
            if (speed != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[speed]')]", XNSMGR).InnerText = boldfont + @"Speed: " + endspan + speed;
            // strength
            if (strength != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[str]')]", XNSMGR).InnerText = strength;
            // dexterity
            if (dexterity != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[dex]')]", XNSMGR).InnerText = dexterity;
            // constitution
            if (constitution != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[con]')]", XNSMGR).InnerText = constitution;
            // intelligence
            if (intelligence != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[int]')]", XNSMGR).InnerText = intelligence;
            // wisdom
            if (wisdom != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[wis]')]", XNSMGR).InnerText = wisdom;
            // charisma
            if (charisma != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[cha]')]", XNSMGR).InnerText = charisma;
            // saving throws
            if (save != null & save != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[saving throws]')]", XNSMGR).InnerText = boldfont + @"Saving Throws: " + endspan + save;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[saving throws]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // skills
            if (skill != null & skill != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[skill]')]", XNSMGR).InnerText = skill;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[skill]')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
            }
            // resist
            if (resist != null & resist != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[resist]')]", XNSMGR).InnerText = boldfont + @"Resistance: " + endspan + resist;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[resist]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // immune
            if (immune != null & immune != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[immune]')]", XNSMGR).InnerText = boldfont + @"Damage Immunities: " + endspan + immune;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[immune]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // condition immunities
            if (conditionimmune != null & conditionimmune != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[condition immunities]')]", XNSMGR).InnerText = boldfont + @"Condition Immunities: " + endspan + conditionimmune;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[condition immunities]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // passive
            if (passive != null & passive != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[passive]')]", XNSMGR).InnerText = boldfont + @"Passive Perception: " + endspan + passive;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[passive]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // senses
            if (senses != null & senses != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[senses]')]", XNSMGR).InnerText = boldfont + @"Senses: " + endspan + senses;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[senses]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // languages
            if (languages != null & languages != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[languages]')]", XNSMGR).InnerText = boldfont + @"Languages: " + endspan + languages;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[languages]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // challenge rating
            if (cr != null & cr != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[cr]')]", XNSMGR).InnerText = boldfont + @"Challenge: " + endspan + cr;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[cr]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // enviornment
            if (enviornment != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[environment]')]", XNSMGR).InnerText = enviornment;
            // spells
            if (spells != null & spells != "")
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[spells]')]", XNSMGR).InnerText = spells;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[spells]')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
                node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), 'Spells')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
            }
            // traits
            if (traits != null)
            {
                XmlNodeList nodelisttrait = Monster.SelectNodes("//trait");
                XmlNode nodetraits = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[traits]')]", XNSMGR).ParentNode.ParentNode;
                MonsterPageXMLDoc = MonsterTableCleanup(MonsterPageXMLDoc, nodelisttrait, nodetraits);
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[traits]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[traits]')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
            }
            // actions
            if (actions != null)
            {
                XmlNodeList nodelistactions = Monster.SelectNodes("//action");
                XmlNode nodeactions = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[actions]')]", XNSMGR).ParentNode.ParentNode;
                MonsterPageXMLDoc = MonsterTableCleanup(MonsterPageXMLDoc, nodelistactions, nodeactions);
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[actions]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[actions]')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
                node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), 'Actions')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
            }
            // legendary
            if (legendary != null & legendary != "")
            {
                XmlNodeList nodelistlegendary = Monster.SelectNodes("//legendary");
                XmlNode nodelegendary = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[legendary]')]", XNSMGR).ParentNode.ParentNode;
                MonsterPageXMLDoc = MonsterTableCleanup(MonsterPageXMLDoc, nodelistlegendary, nodelegendary);
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[legendary]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[legendary]')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
                node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), 'Legendary')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
            }

            try
            {
                MonsterPageXMLDoc.SelectSingleNode("//one:Outline/one:Size/@width", XNSMGR).InnerXml = "400";
                OneNoteApp.UpdatePageContent(MonsterPageXMLDoc.OuterXml);
            }
            catch (Exception ex)
            {
                XDocument.Parse(MonsterPageXMLDoc.OuterXml).Save("errors/" + name + "-onenote.xml");
                XDocument.Parse(Monster.OuterXml).Save("errors/" + name + "-monster.xml");
                Console.WriteLine("failed to update {0}", name);
                Console.WriteLine(ex.Message);
            }


            OneNoteApp.SyncHierarchy(sectionID);

        }

        private void FillSpellTable(string InputXML, string SpellPageID)
        {
            XmlDocument Spell = new XmlDocument();
            Spell.LoadXml(InputXML);

            OneNoteApp.GetHierarchy(SpellPageID, OneNote.HierarchyScope.hsChildren, out string CurrentSpellPage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument SpellPageDocument = new XmlDocument();
            SpellPageDocument.LoadXml(CurrentSpellPage);
            XmlNamespaceManager XNSMGR = XMLMGR(SpellPageDocument);

            // single child nodes
            string name = Spell.SelectSingleNode("//name")?.InnerText;
            string school = Spell.SelectSingleNode("//school")?.InnerText;
            string level = Spell.SelectSingleNode("//level")?.InnerText;
            string time = Spell.SelectSingleNode("//time")?.InnerText;
            string range = Spell.SelectSingleNode("//range")?.InnerText;
            string components = Spell.SelectSingleNode("//components")?.InnerText;
            string duration = Spell.SelectSingleNode("//duration")?.InnerText;
            string description = Spell.SelectSingleNode("//text")?.InnerText;
            string classes = Spell.SelectSingleNode("//classes")?.InnerText;
            string roll = Spell.SelectSingleNode("//roll")?.InnerText;

            school = SpellSchool(school);

            if (level == "0")
                level = "Cantrip";
            else
                level = "Level " + level;

            // name
            if (name != null & name != "")
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = boldname + name + endspan;
            // school
            if (school != null & school != "")
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[school]')]", XNSMGR).InnerText = italicfont + level + ", " + school + endspan;
            // level
            //if (level != null)
            //    SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[level]')]", XNSMGR).InnerText = level;
            // time
            if (time != null & time != "")
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[time]')]", XNSMGR).InnerText = boldfont + "Time to Cast: " + endspan + time;
            // range
            if (range != null & range != "")
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[range]')]", XNSMGR).InnerText = boldfont + "Spell Range: " + endspan + range;
            // components
            if (components != null & components != "")
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[components]')]", XNSMGR).InnerText = boldfont + "Components: " + endspan + components;
            // duration
            if (duration != null & duration != "")
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[duration]')]", XNSMGR).InnerText = boldfont + "Cast Time: " + endspan + duration;
            // description
            if (description != null & description != "")
            {
                string DescriptionText = "";
                XmlNodeList DescriptionElements = Spell.SelectNodes("//text");
                foreach (XmlNode Desc in DescriptionElements)
                {
                    if (Desc.InnerText != "")
                    {

                        DescriptionText += Desc.InnerText + Environment.NewLine + Environment.NewLine;
                    }

                }

                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[description]')]", XNSMGR).InnerText = DescriptionText;
            }
            else
            {
                XmlNode node = SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[description]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }

            // classes
            if (classes != null & classes != "")
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[classes]')]", XNSMGR).InnerText = classes;

            XmlNode rollnode = SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[roll]')]", XNSMGR);
            rollnode.ParentNode.ParentNode.RemoveChild(rollnode.ParentNode);

            try
            {
                SpellPageDocument.SelectSingleNode("//one:Outline/one:Size/@width", XNSMGR).InnerXml = "400";
                OneNoteApp.UpdatePageContent(SpellPageDocument.OuterXml);
            }
            catch (Exception ex)
            {
                XDocument.Parse(SpellPageDocument.OuterXml).Save("errors/" + name + "-onenote.xml");
                XDocument.Parse(Spell.OuterXml).Save("errors/" + name + "-spells.xml");
                Console.WriteLine("failed to update {0}", name);
                Console.WriteLine(ex.Message);
            }


            OneNoteApp.SyncHierarchy(sectionID);

        }

        private void FillRaceTable(string InputXML, string RacePageID)
        {
            XmlDocument Race = new XmlDocument();
            Race.LoadXml(InputXML);

            OneNoteApp.GetHierarchy(RacePageID, OneNote.HierarchyScope.hsChildren, out string CurrentRacePage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument RacePageDocument = new XmlDocument();
            RacePageDocument.LoadXml(CurrentRacePage);
            XmlNamespaceManager XNSMGR = XMLMGR(RacePageDocument);

            // single child nodes
            string name = Race.SelectSingleNode("//name")?.InnerText;
            string size = MonsterSize(Race.SelectSingleNode("//size")?.InnerText.ToUpper());
            string speed = Race.SelectSingleNode("//speed")?.InnerText;
            string ability = Race.SelectSingleNode("//ability")?.InnerText;
            string trait = Race.SelectSingleNode("//trait")?.InnerText;
            string proficiency = Race.SelectSingleNode("//proficiency")?.InnerText;

            // name
            if (name != null)
                RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = boldname + name + endspan;

            if (size != null & size != "")
                RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[size]')]", XNSMGR).InnerText = boldfont + "Size: " + endspan + size;
            else
            {
                XmlNode node = RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[size]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }

            if (speed != null & speed != "")
                RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[speed]')]", XNSMGR).InnerText = boldfont + "Speed: " + endspan + speed;
            else
            {
                XmlNode node = RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[speed]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            if (ability != null | proficiency != null)
            {
                if (ability != null & ability != "")
                    RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[ability]')]", XNSMGR).InnerText = boldfont + "Abilities: " + endspan + ability;
                else
                {
                    XmlNode node = RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[ability]')]", XNSMGR);
                    node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
                }

                if (proficiency != null)
                    RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[proficiency]')]", XNSMGR).InnerText = boldfont + "Proficiencies: " + endspan + proficiency;
                else
                {
                    XmlNode node = RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[proficiency]')]", XNSMGR);
                    node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
                }
            }
            else
            {
                XmlNode node = RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[ability]')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
            }


            if (trait != null)
            {
                string objstring = "";
                XmlNodeList elementlist = Race.SelectNodes("//trait");
                foreach (XmlNode node in elementlist)
                {
                    if (node.InnerText != "")
                    {
                        foreach (XmlNode childnode in node.ChildNodes)
                        {
                            if (childnode.InnerText != "")
                            {
                                if (childnode.Name == "name")
                                    objstring += boldfont + childnode.InnerText + "." + endspan + Environment.NewLine;
                                else if (childnode.Name == "text")
                                    objstring += childnode.InnerText + Environment.NewLine + Environment.NewLine;
                            }
                        }
                    }
                }

                RacePageDocument.SelectSingleNode("//one:T[contains(text(), '[traits]')]", XNSMGR).InnerText = objstring;
            }


            try
            {
                RacePageDocument.SelectSingleNode("//one:Outline/one:Size/@width", XNSMGR).InnerXml = "400";
                OneNoteApp.UpdatePageContent(RacePageDocument.OuterXml);
            }
            catch (Exception ex)
            {
                //XDocument.Parse(RacePageDocument.OuterXml).Save("errors/" + name + "-onenote.xml");
                //XDocument.Parse(Race.OuterXml).Save("errors/" + name + "-race.xml");
                Console.WriteLine("failed to update {0}", name);
                Console.WriteLine(ex.Message);
            }




            OneNoteApp.SyncHierarchy(sectionID);

        }

        private void FillFeatTable(string InputXML, string FeatPageID)
        {

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(InputXML);

            //Console.WriteLine("");
            //Console.WriteLine(xmldoc.OuterXml);
            //Console.ReadKey(true);

            OneNoteApp.GetHierarchy(FeatPageID, OneNote.HierarchyScope.hsChildren, out string CurrentFeatPage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument FeatPageDocument = new XmlDocument();
            FeatPageDocument.LoadXml(CurrentFeatPage);
            XmlNamespaceManager XNSMGR = XMLMGR(FeatPageDocument);

            // single child nodes
            string name = xmldoc.SelectSingleNode("//name")?.InnerText;
            string modifier = xmldoc.SelectSingleNode("//modifier")?.InnerText;
            string prerequisites = xmldoc.SelectSingleNode("//prerequisite")?.InnerText;
            string description = xmldoc.SelectSingleNode("//text")?.InnerText;

            Console.WriteLine("{0}name = {1}{0}modifier = {2}{0}prerequisites = {3}{0}description = {4}{0}", Environment.NewLine, name, modifier, prerequisites, description);
            //Console.ReadKey(true);

            if (name.Contains("(UA)"))
                name = name.Replace("(UA)", "").Trim();

            // name
            if (name != null & name != "")
                FeatPageDocument.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = boldname + name + endspan;

            if (modifier != null & modifier != "")
                FeatPageDocument.SelectSingleNode("//one:T[contains(text(), '[modifier]')]", XNSMGR).InnerText = bolditalicfont + "Modifier/s: " + endspan + italicfont + modifier + endspan;
            else
            {
                XmlNode node = FeatPageDocument.SelectSingleNode("//one:T[contains(text(), '[modifier]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }

            if (prerequisites != null & prerequisites != "")
                FeatPageDocument.SelectSingleNode("//one:T[contains(text(), '[prerequisites]')]", XNSMGR).InnerText = boldfont + "Prerequisites: " + endspan + prerequisites;
            else
            {
                XmlNode node = FeatPageDocument.SelectSingleNode("//one:T[contains(text(), '[prerequisites]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }

            if (description.Length > 0)
            {
                string objstring = "";
                XmlNodeList elementlist = xmldoc.SelectNodes("//text");
                foreach (XmlNode node in elementlist)
                {
                    if (node.InnerText != "")
                    {
                        foreach (XmlNode childnode in node.ChildNodes)
                        {
                            if (childnode.InnerText != "")
                            {
                                if (childnode.InnerText.Contains("•"))
                                    objstring += "   " + childnode.InnerText + Environment.NewLine;
                                else
                                    objstring += childnode.InnerText + Environment.NewLine;
                            }
                        }
                    }
                }

                FeatPageDocument.SelectSingleNode("//one:T[contains(text(), '[description]')]", XNSMGR).InnerText = objstring;
            }
            else
            {
                XmlNode node = FeatPageDocument.SelectSingleNode("//one:T[contains(text(), '[description]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }

            try
            {
                FeatPageDocument.SelectSingleNode("//one:Outline/one:Size/@width", XNSMGR).InnerXml = "400";
                OneNoteApp.UpdatePageContent(FeatPageDocument.OuterXml);
            }
            catch (Exception ex)
            {
                //XDocument.Parse(RacePageDocument.OuterXml).Save("errors/" + name + "-onenote.xml");
                //XDocument.Parse(Race.OuterXml).Save("errors/" + name + "-race.xml");
                Console.WriteLine("failed to update {0}", name);
                Console.WriteLine(ex.Message);
            }


            OneNoteApp.SyncHierarchy(sectionID);

        }

        private void FillBackgroundTable(string InputXML, string PageID)
        {
            XmlDocument InputNode = new XmlDocument();
            InputNode.LoadXml(InputXML);

            OneNoteApp.GetHierarchy(PageID, OneNote.HierarchyScope.hsChildren, out string CurrentPage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(CurrentPage);
            XmlNamespaceManager XNSMGR = XMLMGR(xmldoc);

            // single child nodes
            string name = InputNode.SelectSingleNode("//name")?.InnerText;
            string trait = InputNode.SelectSingleNode("//trait")?.InnerText;
            //string proficiency = InputNode.SelectSingleNode("proficiency")?.InnerText;

            // name
            if (name != null)
                xmldoc.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = boldname + name + endspan;

            if (trait != null)
            {
                string objstring = "";
                XmlNodeList elementlist = InputNode.SelectNodes("//trait");
                foreach (XmlNode node in elementlist)
                {
                    if (node.InnerText != "")
                    {
                        foreach (XmlNode childnode in node.ChildNodes)
                        {
                            if (childnode.InnerText != "")
                            {
                                if (childnode.Name == "name")
                                    objstring += boldfont + childnode.InnerText + "." + endspan + Environment.NewLine;
                                else if (childnode.Name == "text")
                                    objstring += childnode.InnerText + Environment.NewLine + Environment.NewLine;
                            }
                        }
                    }
                }

                xmldoc.SelectSingleNode("//one:T[contains(text(), '[traits]')]", XNSMGR).InnerText = objstring;
            }
            //if (proficiency != null)
            //    xmldoc.SelectSingleNode("//one:T[contains(text(), '[proficiency]')]", XNSMGR).InnerText = bold + "Proficiencies: " + endspan + proficiency;
            //else
            //{
            //    XmlNode node = xmldoc.SelectSingleNode("//one:T[contains(text(), '[proficiency]')]", XNSMGR);
            //}

            try
            {
                xmldoc.SelectSingleNode("//one:Outline/one:Size/@width", XNSMGR).InnerXml = "400";
                OneNoteApp.UpdatePageContent(xmldoc.OuterXml);
            }
            catch (Exception ex)
            {
                //XDocument.Parse(RacePageDocument.OuterXml).Save("errors/" + name + "-onenote.xml");
                //XDocument.Parse(Race.OuterXml).Save("errors/" + name + "-race.xml");
                Console.WriteLine("failed to update {0}", name);
                Console.WriteLine(ex.Message);
            }


            OneNoteApp.SyncHierarchy(sectionID);

        }

        private void FillClassTable(string InputXML, string PageID)
        {
            XmlDocument InputNode = new XmlDocument();
            InputNode.LoadXml(InputXML);

            OneNoteApp.GetHierarchy(PageID, OneNote.HierarchyScope.hsChildren, out string CurrentPage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(CurrentPage);
            XmlNamespaceManager XNSMGR = XMLMGR(xmldoc);

            string name = InputNode.SelectSingleNode("name")?.InnerText;
            string hd = InputNode.SelectSingleNode("hd")?.InnerText;
            string proficiency = InputNode.SelectSingleNode("proficiency")?.InnerText;
            string spellAbility = InputNode.SelectSingleNode("spellAbility")?.InnerText;
            string slots = InputNode.SelectSingleNode("slots")?.InnerText;
            string feature = InputNode.SelectSingleNode("feature")?.InnerText;
            string spellslots = InputNode.SelectSingleNode("//class[name[text()='" + name + "']]//slots")?.InnerText;

            string HitDice = string.Format("{0} Hit Dice: {1}1d{2} per {3} level", boldfont, endspan, hd, name.ToLower());
            string HitPoints = string.Format("{0} Hit Points at 1st Level: {1}{2} + your constitution modifier", boldfont, endspan, hd);
            string HitPointsFuture = string.Format("{0} Hit Points at Higher Levels: {1}1d{2} + your Constitution modifier per {3} level after 1st", boldfont, endspan, hd, name.ToLower());

            Console.WriteLine("current class = {0}", name);

            if (name != null)
                xmldoc.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = boldname + name + endspan;
            if (hd != null)
                xmldoc.SelectSingleNode("//one:T[contains(text(), '[hd]')]", XNSMGR).InnerText = HitDice + Environment.NewLine + HitPoints + Environment.NewLine + HitPointsFuture;
            if (proficiency != null)
                xmldoc.SelectSingleNode("//one:T[contains(text(), '[proficiency]')]", XNSMGR).InnerText = boldfont + "Proficiencies: " + endspan + proficiency;
            else
            {
                XmlNode node = xmldoc.SelectSingleNode("//one:T[contains(text(), '[proficiency]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }

            if (spellslots != null | spellslots != "")
            {
                string spellSlotsXPATH = string.Format("//class[name[contains(text(), '{0}')]]//autolevel[slots]", name);
                SpellSlotTable(xmldoc, InputNode.SelectNodes(spellSlotsXPATH, XNSMGR), pageID);
            }
            else
            {
                XmlNode node = xmldoc.SelectSingleNode("//one:T[contains(text(), '[spellslots]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }

            if (spellAbility != null & spellAbility != "")
                xmldoc.SelectSingleNode("//one:T[contains(text(), '[spellability]')]", XNSMGR).InnerText = boldfont + "Spell Ability: " + endspan + spellAbility;
            else
            {
                XmlNode node = xmldoc.SelectSingleNode("//one:T[contains(text(), '[spellability]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }

            try
            {
                XDocument.Parse(xmldoc.OuterXml).Save("testfile.xml");
                xmldoc.SelectSingleNode("//one:Outline/one:Size/@width", XNSMGR).InnerXml = "400";
                OneNoteApp.UpdatePageContent(xmldoc.OuterXml);
            }
            catch (Exception ex)
            {
                Console.WriteLine("failed to update {0}", name);
                Console.WriteLine(ex.Message);
            }

            OneNoteApp.SyncHierarchy(sectionID);

            //Console.ReadKey(true);

        }

        // Functions

        XmlNamespaceManager XMLMGR(XmlDocument InputXMLDocument)
        {
            XmlNamespaceManager XMLMGR = new XmlNamespaceManager(InputXMLDocument.NameTable);
            XMLMGR.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");
            return XMLMGR;
        }

        private string ID(string XMLFILE, string inputAttribute)
        {
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(XMLFILE);
            string XPATH = string.Format("//*[@name='{0}']/@ID", inputAttribute);

            try
            {
                return xmldoc.SelectSingleNode(XPATH, XMLMGR(xmldoc)).Value;
            }
            catch (Exception)
            {
                return "";
            }


        }

        private XmlDocument MonsterTableCleanup(XmlDocument inputDocument, XmlNodeList nodeList, XmlNode onenoteNodeToUpdate)
        {
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            foreach (XmlNode statusTrait in nodeList)
            {
                string traits = "";
                string attack = "";

                XmlNode oneOE = inputDocument.CreateElement("one:OE", URI);
                XmlNode oneT = inputDocument.CreateElement("one:T", URI);

                if (statusTrait.SelectSingleNode("attack") != null)
                {
                    string nodetext = statusTrait.SelectSingleNode("attack").InnerText;
                    //traits += font + nodetext + endspan + Environment.NewLine;
                    try
                    {
                        if (nodetext.Split('|')[1] != "")
                            attack = " (" + nodetext.Split('|')[1] + " | " + nodetext.Split('|')[2] + ")";
                        else
                            attack = " (" + nodetext.Split('|')[2] + ")";
                    }
                    catch
                    {
                        attack = "";
                    }
                }

                // add the monster name
                traits += boldname + statusTrait.SelectSingleNode("name").InnerText + endspan + font + attack + endspan + Environment.NewLine;

                // get list of <text> elements form monster bestiery
                XmlNodeList text = statusTrait.SelectNodes("text");

                foreach (XmlNode nodetraittext in text)
                {
                    if (nodetraittext.InnerText != "")
                    {
                        string nodetext = nodetraittext.InnerText;

                        if (nodetext.Contains("•"))
                            nodetext = nodetext.Replace("•", "    • ");

                        traits += font + nodetext + endspan + Environment.NewLine;
                    }

                }



                oneT.InnerText = traits;
                oneOE.AppendChild(oneT);
                onenoteNodeToUpdate.AppendChild(oneOE);

            }

            return inputDocument;
        }

        private string MonsterSize(string inputSize)
        {
            if (inputSize == null | inputSize == "")
                return null;

            inputSize = inputSize.Replace("S", "Small");
            inputSize = inputSize.Replace("M", "Medium");
            inputSize = inputSize.Replace("L", "Large");
            inputSize = inputSize.Replace("H", "Huge");
            inputSize = inputSize.Replace("G", "Gargantuan");
            return inputSize;
        }

        private string ChallengeRating(string inputCR)
        {
            if (inputCR == null)
                return null;

            if (inputCR == "0")
                return "0 (0 or 10 xp)";
            if (inputCR == "1/8")
                return "1/8 (25 xp)";
            if (inputCR == "1/4")
                return "1/4 (50 xp)";
            if (inputCR == "1/2")
                return "1/2 (100 xp)";
            if (inputCR == "1")
                return "1 (200 xp)";
            if (inputCR == "2")
                return "2 (450 xp)";
            if (inputCR == "3")
                return "3 (700 xp)";
            if (inputCR == "4")
                return "4 (1,100 xp)";
            if (inputCR == "5")
                return "5 (1,800 xp)";
            if (inputCR == "6")
                return "6 (2,300 xp)";
            if (inputCR == "7")
                return "7 (2,900 xp)";
            if (inputCR == "8")
                return "8 (3,900 xp)";
            if (inputCR == "9")
                return "9 (5,000 xp)";
            if (inputCR == "10")
                return "10 (5,900 xp)";
            if (inputCR == "11")
                return "11 (7,200 xp)";
            if (inputCR == "12")
                return "12 (8,400 xp)";
            if (inputCR == "13")
                return "13 (10,000 xp)";
            if (inputCR == "14")
                return "14 (11,500 xp)";
            if (inputCR == "15")
                return "15 (13,000 xp)";
            if (inputCR == "16")
                return "16 (15,000 xp)";
            if (inputCR == "17")
                return "17 (18,000 xp)";
            if (inputCR == "18")
                return "18 (20,000 xp)";
            if (inputCR == "19")
                return "19 (22,000 xp)";
            if (inputCR == "20")
                return "20 (25,000 xp)";
            if (inputCR == "21")
                return "21 (33,000 xp)";
            if (inputCR == "22")
                return "22 (41,000 xp)";
            if (inputCR == "23")
                return "23 (50,000 xp)";
            if (inputCR == "24")
                return "24 (62,000 xp)";
            if (inputCR == "25")
                return "25 (75,000 xp)";
            if (inputCR == "26")
                return "26 (90,000 xp)";
            if (inputCR == "27")
                return "27 (105,000 xp)";
            if (inputCR == "28")
                return "28 (120,000 xp)";
            if (inputCR == "29")
                return "29 (135,000 xp)";
            if (inputCR == "30")
                return "30 (155,000 xp)";

            return "";

        }

        private string SpellSchool(string inputSize)
        {
            inputSize = inputSize.Replace("C", "Conjuration");
            inputSize = inputSize.Replace("A", "Abjuration");
            inputSize = inputSize.Replace("T", "Transmutation");
            inputSize = inputSize.Replace("I", "Illusion");
            inputSize = inputSize.Replace("EV", "Evocation");
            inputSize = inputSize.Replace("EN", "Enchantment");
            inputSize = inputSize.Replace("D", "Divination");
            inputSize = inputSize.Replace("N", "Necromancy");
            inputSize = inputSize.Replace("U", "Universal");
            return inputSize;
        }

        private XmlNode SpellSlotTable(XmlDocument inputXMLDocument, XmlNodeList SpellSlotLevels, string pageID)
        {

            if (SpellSlotLevels.Count > 0)
            {
                Console.WriteLine("building spellslot table");
                string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";
                XmlNamespaceManager XNSMGR = XMLMGR(inputXMLDocument);
                string name = inputXMLDocument.SelectSingleNode("//one:Title", XNSMGR).InnerText;

                string xpathclassbyname = string.Format("//class[name[text()='{0}']]", name);

                XmlNode Parent = inputXMLDocument.SelectSingleNode("//one:T[contains(text(), '[spellslots]')]/parent::one:OE/parent::one:OEChildren", XNSMGR);
                XmlNode Table = inputXMLDocument.CreateElement("one:Table", URI);
                XmlNode Columns = inputXMLDocument.CreateElement("one:Columns", URI);
                XmlNode Column = null;
                XmlNode Row = inputXMLDocument.CreateElement("one:Row", URI);
                XmlNode Cell = null;
                XmlNode OEChildren = null;
                XmlNode OE = inputXMLDocument.CreateElement("one:OE", URI);
                XmlNode T = inputXMLDocument.CreateElement("one:T", URI);
                XmlNode OETable = inputXMLDocument.CreateElement("one:OE", URI);

                XmlAttribute TableBorders = inputXMLDocument.CreateAttribute("bordersVisible");
                TableBorders.Value = "true";
                Table.Attributes.Append(TableBorders);

                XmlNode slottext = SpellSlotLevels.Item(0).SelectSingleNode(xpathclassbyname + "/autolevel/slots");
                string[] SpellSlots = slottext.InnerText.Split(',');
                for (int i = 0; i < SpellSlots.Length; i++)
                {
                    Column = inputXMLDocument.CreateElement("one:Column", URI);
                    XmlAttribute ColumnIndex = inputXMLDocument.CreateAttribute("index");
                    XmlAttribute ColumnWidth = inputXMLDocument.CreateAttribute("width");
                    ColumnIndex.Value = i.ToString();
                    ColumnWidth.Value = "37.1100006103516";
                    Column.Attributes.Append(ColumnIndex);
                    Column.Attributes.Append(ColumnWidth);
                    Columns.AppendChild(Column);
                }

                Table.AppendChild(Columns);

                // add headers
                List<string> headers = new List<string> { "Level", "Cantrips", "Slot 1", "Slot 2", "Slot 3", "Slot 4", "Slot 5", "Slot 6", "Slot 7", "Slot 8", "Slot 9" };
                Row = inputXMLDocument.CreateElement("one:Row", URI);
                int columnnumber = 0;
                foreach (string header in headers)
                {
                    if (columnnumber <= SpellSlots.Length)
                    {
                        Cell = inputXMLDocument.CreateElement("one:Cell", URI);
                        OEChildren = inputXMLDocument.CreateElement("one:OEChildren", URI);
                        OE = inputXMLDocument.CreateElement("one:OE", URI);
                        T = inputXMLDocument.CreateElement("one:T", URI);

                        XmlAttribute CellShading = inputXMLDocument.CreateAttribute("shadingColor");
                        CellShading.Value = "#000000";
                        Cell.Attributes.Append(CellShading);

                        T.InnerText = tableheaderblack + header + endspan;

                        OE.AppendChild(T);
                        OEChildren.AppendChild(OE);
                        Cell.AppendChild(OEChildren);
                        Row.AppendChild(Cell);
                    }
                    columnnumber++;
                }

                Table.AppendChild(Row);

                int level = int.Parse(slottext.ParentNode.Attributes.GetNamedItem("level").Value);
                Console.WriteLine("first level for {0} to get some spells is {1}", name, level);
                foreach (XmlNode Level in SpellSlotLevels)
                {
                    Row = inputXMLDocument.CreateElement("one:Row", URI);

                    slottext = Level.SelectSingleNode(xpathclassbyname + "//slots");


                    Cell = inputXMLDocument.CreateElement("one:Cell", URI);
                    XmlAttribute CellShading = inputXMLDocument.CreateAttribute("shadingColor");
                    CellShading.Value = "#FFFFFF";
                    Cell.Attributes.Append(CellShading);
                    OEChildren = inputXMLDocument.CreateElement("one:OEChildren", URI);
                    OE = inputXMLDocument.CreateElement("one:OE", URI);
                    T = inputXMLDocument.CreateElement("one:T", URI);
                    T.InnerText = tablefont + level.ToString() + endspan;
                    OE.AppendChild(T);
                    OEChildren.AppendChild(OE);
                    Cell.AppendChild(OEChildren);
                    Row.AppendChild(Cell);

                    SpellSlots = slottext.InnerText.Split(',');

                    foreach (string SpellSlot in SpellSlots)
                    {
                        string SpellSlotText = SpellSlot;

                        if (SpellSlotText == "0")
                            SpellSlotText = "-";

                        Cell = inputXMLDocument.CreateElement("one:Cell", URI);
                        CellShading = inputXMLDocument.CreateAttribute("shadingColor");
                        CellShading.Value = "#FFFFFF";
                        Cell.Attributes.Append(CellShading);
                        OEChildren = inputXMLDocument.CreateElement("one:OEChildren", URI);
                        OE = inputXMLDocument.CreateElement("one:OE", URI);
                        T = inputXMLDocument.CreateElement("one:T", URI);
                        T.InnerText = tablefont + SpellSlotText + endspan;
                        OE.AppendChild(T);
                        OEChildren.AppendChild(OE);
                        Cell.AppendChild(OEChildren);
                        Row.AppendChild(Cell);
                    }
                    level++;
                    Table.AppendChild(Row);
                }

                OETable.AppendChild(Table);
                Parent.InsertBefore(OETable, Parent.FirstChild);
                Parent.RemoveChild(Parent.SelectSingleNode("//one:T[contains(text(), '[spellslots]')]/parent::one:OE", XNSMGR));
                XDocument.Parse(inputXMLDocument.OuterXml).Save("newtableforspellslots.xml");

                OneNoteApp.UpdatePageContent(inputXMLDocument.OuterXml);
                OneNoteApp.SyncHierarchy(pageID);
            }
            return null;
        }

        private XElement NewTable(int numColums, int numRows, string hexHeaderShading, string hexCellShading)
        {

            string columns = "";
            string cells = "";
            string headercells = "";
            string rows = "";

            for (int columnindex = 0; columnindex < numColums; columnindex++)
            {
                columns += string.Format(@"<one:Column index='{0}' width='37' />", columnindex);
                cells += @"<one:Cell objectID=''><one:OEChildren><one:OE objectID=''><one:T></one:T></one:OE></one:OEChildren></one:Cell>";
                headercells += @"<one:Cell objectID=''><one:OEChildren><one:OE objectID=''><one:T></one:T></one:OE></one:OEChildren></one:Cell>";
            }

            for (int rowindex = 0; rowindex < numRows; rowindex++)
            {
                rows += string.Format(@"<one:Row objectID=''>{0}</one:Row>", cells);
            }

            string strOneNoteTableXML =
                    "<one:OE objectID='' xmlns:one='http://schemas.microsoft.com/office/onenote/2010/onenote'>>" +
                        "<one:Table objectID=''>" +
                            "<one:Columns>" +
                                columns +
                            "</one:Columns>" +
                            "<one:Row>" +
                                headercells +
                            "</one:Row>" +
                            rows +
                        "</one:Table>" +
                    "</one:OE>";

            return XElement.Parse(strOneNoteTableXML);
        }

        private XElement NewOutline(XElement descendant, double x = 36.0, double y = 68.4000015258789, int z = 1, double width = 120, double height = 13.4)
        {
            string strOneNoteXMLFormatted =
                string.Format(
                        "<Outline>" +
                            "<Position x='{0}' y='{1}' z='{2}' />" +
                            "<Size width = '{3}' height = '{4}' />" +
                            "<OEChildren>" +
                                "<OE>" +
                                    "{5}" +
                                "</OE>" +
                            "</OEChildren>" +
                        "</Outline>", x, y, z, width, height, "<T>test</T>");
            //String.Concat(descendant.Nodes())
            return XElement.Parse(strOneNoteXMLFormatted);
        }
    }
}
