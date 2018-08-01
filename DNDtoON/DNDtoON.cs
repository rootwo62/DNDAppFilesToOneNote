using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OneNote = Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.Data;
using System.IO;
using System.Xml.Linq;


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
        const string italicfont = @"<span style='font-style:italic; font-size:9pt; font-family:cambria'>";
        const string bolditalicfont = @"<span style='font-style:italic; font-weight:bold; font-size:9pt; font-family:cambria'>";
        const string endspan = @"</span>";

        static void Main(string[] args)
        {

            Application Application = new Application();
            Application.Run();

            Console.WriteLine("press any key to close...");
            Console.ReadKey(true);
        }

        private void Run()
        {

            string ElementsByName = Properties.Settings.Default.ElementsByName;
            List<string> listofnames = new List<string> { };
            string blocktype = Properties.Settings.Default.BlockType;
            string templatexmlfile = blocktype + "templatefile.xml";
            GetOneNoteTableXML(Properties.Settings.Default.BlockTemplatePageName, blocktype + "templatefile.xml");
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(Properties.Settings.Default.DNDAppFileXML);
            XmlNodeList docnodes = null;

            if (ElementsByName.StartsWith("[list]"))
            {
                Console.WriteLine("getting elements by name...");
                try
                {
                    foreach (string name in ElementsByName.Split(','))
                    {
                        string nametoadd = name.Replace("[list]", "").Trim();
                        listofnames.Add(nametoadd);
                        Console.WriteLine("added {0} to the list of names", nametoadd);
                    }
                }
                catch (Exception)
                {
                    throw;
                }

            }

            if (listofnames.Count > 0)
            {
                XmlDocument listnodes = new XmlDocument();
                XmlElement nodes = listnodes.CreateElement(blocktype + "s");

                foreach (string item in listofnames)
                {
                    XmlNode importnode = listnodes.ImportNode(xmldoc.SelectSingleNode("//" + blocktype + "[name[text()='" + item + "']]"), true);
                    nodes.AppendChild(importnode);
                }

                docnodes = listnodes.AppendChild(nodes).ChildNodes;
            }
            else
            {
                docnodes = xmldoc.DocumentElement.GetElementsByTagName("race");
            }

                foreach (XmlNode node in docnodes)
                {
                    string pageName = node.SelectSingleNode("name").InnerText;
                    Console.WriteLine("adding {0} to section {1} in notebook {2}", pageName, sectionname, notebookName);
                    CopyPageTableTemplate(pageName, templatexmlfile, out string pageID);

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

                }
        }

        private void GetONHeierarchy(string ID, OneNote.HierarchyScope OneNoteHierarchyScope, string OutputFile)
        {
            OneNoteApp.GetHierarchy(ID, OneNoteHierarchyScope, out string pbstrHierarchyXMLOut);
            XDocument.Parse(pbstrHierarchyXMLOut).Save(OutputFile);
        }

        private void InsertElementIntoONPage(string PageXML, XmlNode NodeToAppend)
        {
            XmlDocument oldPageXmlDocument = new XmlDocument();
            oldPageXmlDocument.LoadXml(PageXML);
            XmlNode importNode = oldPageXmlDocument.ImportNode(NodeToAppend, true);
            oldPageXmlDocument.SelectSingleNode("//one:Outline//OE", XMLMGR(oldPageXmlDocument)).AppendChild(NodeToAppend);
            OneNoteApp.UpdatePageContent(oldPageXmlDocument.OuterXml);
            OneNoteApp.SyncHierarchy(sectionID);
        }

        private void CreateONPage(string NewPageName, string ONHierarchyID, out string NewPageID, out string NewPageXML)
        {
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            OneNoteApp.CreateNewPage(ONHierarchyID, out NewPageID);
            OneNoteApp.GetPageContent(NewPageID, out NewPageXML);

            XmlDocument NewPageXmlDocument = new XmlDocument();
            NewPageXmlDocument.LoadXml(NewPageXML);
            XmlNode PageXML = NewPageXmlDocument.SelectSingleNode("//one:Page", XMLMGR(NewPageXmlDocument));

            XmlNode Outline = NewPageXmlDocument.CreateElement("one:Outline", URI);
            XmlNode Position = NewPageXmlDocument.CreateElement("one:Position", URI);
            XmlNode Size = NewPageXmlDocument.CreateElement("one:Size", URI);
            XmlNode OEChildren = NewPageXmlDocument.CreateElement("one:OEChildren", URI);
            XmlNode OE = NewPageXmlDocument.CreateElement("one:OE", URI);

            XmlAttribute OutlineXMLNS = NewPageXmlDocument.CreateAttribute("xmlns:one");

            OutlineXMLNS.Value = URI;

            Outline.Attributes.Append(OutlineXMLNS);

            XmlAttribute OutlinePositionx = NewPageXmlDocument.CreateAttribute("x");
            XmlAttribute OutlinePositiony = NewPageXmlDocument.CreateAttribute("y");
            XmlAttribute OutlinePositionz = NewPageXmlDocument.CreateAttribute("z");

            OutlinePositionx.Value = "36.0";
            OutlinePositiony.Value = "86.4000015258789";
            OutlinePositionz.Value = "0";

            Position.Attributes.Append(OutlinePositionx);
            Position.Attributes.Append(OutlinePositiony);
            Position.Attributes.Append(OutlinePositionz);

            XmlAttribute OutlineSizeWidth = NewPageXmlDocument.CreateAttribute("width");
            XmlAttribute OutlineSizeHeight = NewPageXmlDocument.CreateAttribute("height");

            OutlineSizeWidth.Value = "260.3399963378906";
            OutlineSizeHeight.Value = "347.3149719238281";

            Size.Attributes.Append(OutlineSizeHeight);
            Size.Attributes.Append(OutlineSizeWidth);

            OEChildren.AppendChild(OE);
            PageXML.AppendChild(Outline);
            Outline.AppendChild(Position);
            Outline.AppendChild(Size);
            Outline.AppendChild(OEChildren);

            NewPageXmlDocument.SelectSingleNode("//one:Title//one:T", XMLMGR(NewPageXmlDocument)).InnerText = NewPageName;

            OneNoteApp.UpdatePageContent(NewPageXmlDocument.OuterXml);
            OneNoteApp.SyncHierarchy(sectionID);
        }

        // Load Variables For Specified Block

        private void GetOneNoteTableXML(string templatePageName, string outXMLFile)
        {
            sectionname = Properties.Settings.Default.OneNoteSection;
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
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            OneNoteApp.CreateNewPage(sectionID, out pageID);
            OneNoteApp.GetPageContent(pageID, out string NewPageXML);

            //XDocument.Parse(NewPageXML).Save("test2.xml");

            XmlDocument XMLDOCUMENT = new XmlDocument();
            XMLDOCUMENT.Load(templatePageXML);
            XmlNode TableElement = XMLDOCUMENT.SelectSingleNode("//one:Table", XMLMGR(XMLDOCUMENT));

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(NewPageXML);
            XmlNode PageXML = xmldoc.SelectSingleNode("//one:Page", XMLMGR(xmldoc));

            XmlNode Outline = xmldoc.CreateElement("one:Outline", URI);
            XmlNode Position = xmldoc.CreateElement("one:Position", URI);
            XmlNode Size = xmldoc.CreateElement("one:Size", URI);
            XmlNode OEChildren = xmldoc.CreateElement("one:OEChildren", URI);
            XmlNode OE = xmldoc.CreateElement("one:OE", URI);

            //XmlNode OEChildren = xmldoc.CreateElement("one:Outline", URI);

            XmlAttribute OutlineXMLNS = xmldoc.CreateAttribute("xmlns:one");

            OutlineXMLNS.Value = URI;

            Outline.Attributes.Append(OutlineXMLNS);

            XmlAttribute OutlinePositionx = xmldoc.CreateAttribute("x");
            XmlAttribute OutlinePositiony = xmldoc.CreateAttribute("y");
            XmlAttribute OutlinePositionz = xmldoc.CreateAttribute("z");

            OutlinePositionx.Value = "36.0";
            OutlinePositiony.Value = "86.4000015258789";
            OutlinePositionz.Value = "0";

            Position.Attributes.Append(OutlinePositionx);
            Position.Attributes.Append(OutlinePositiony);
            Position.Attributes.Append(OutlinePositionz);

            XmlAttribute OutlineSizeWidth = xmldoc.CreateAttribute("width");
            XmlAttribute OutlineSizeHeight = xmldoc.CreateAttribute("height");

            OutlineSizeWidth.Value = "260.3399963378906";
            OutlineSizeHeight.Value = "347.3149719238281";

            Size.Attributes.Append(OutlineSizeHeight);
            Size.Attributes.Append(OutlineSizeWidth);

            XmlNode importNode = xmldoc.ImportNode(TableElement, true);

            OE.AppendChild(importNode);
            OEChildren.AppendChild(OE);
            PageXML.AppendChild(Outline);
            Outline.AppendChild(Position);
            Outline.AppendChild(Size);
            Outline.AppendChild(OEChildren);

            xmldoc.SelectSingleNode("//one:Title//one:T", XMLMGR(xmldoc)).InnerText = outPageName;

            //PageXML.AppendChild(importNode);

            //xmldoc.SelectSingleNode("//one:Outline/@objectID", XMLMGR(xmldoc)).InnerXml = "{" + Guid.NewGuid().ToString().ToUpper() + "}";

            //xmldoc.Save("test.xml");
            //Console.WriteLine(xmldoc.OuterXml);
            //OneNoteApp.UpdateHierarchy(xmldoc.OuterXml);
            OneNoteApp.UpdatePageContent(xmldoc.OuterXml);
            OneNoteApp.SyncHierarchy(sectionID);

        }

        // Fill Tables

        
        private void FillMonsterStatsTable(XmlNode Monster, string MonsterPageID)
        {

            OneNoteApp.GetHierarchy(MonsterPageID, OneNote.HierarchyScope.hsChildren, out string CurrentMonsterPageXml);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument MonsterPageXMLDoc = new XmlDocument();
            MonsterPageXMLDoc.LoadXml(CurrentMonsterPageXml);
            XmlNamespaceManager XNSMGR = XMLMGR(MonsterPageXMLDoc);
            XmlNode Table = MonsterPageXMLDoc.SelectSingleNode("//one:Table", XNSMGR);

            // single child nodes
            string name = Monster.SelectSingleNode("name")?.InnerText;
            string size = Monster.SelectSingleNode("size")?.InnerText;
            string type = Monster.SelectSingleNode("type")?.InnerText;
            string alignment = Monster.SelectSingleNode("alignment")?.InnerText;
            string ac = Monster.SelectSingleNode("ac")?.InnerText;
            string hp = Monster.SelectSingleNode("hp")?.InnerText;
            string speed = Monster.SelectSingleNode("speed")?.InnerText;
            string strength = Monster.SelectSingleNode("str")?.InnerText;
            string dexterity = Monster.SelectSingleNode("dex")?.InnerText;
            string constitution = Monster.SelectSingleNode("con")?.InnerText;
            string intelligence = Monster.SelectSingleNode("int")?.InnerText;
            string wisdom = Monster.SelectSingleNode("wis")?.InnerText;
            string charisma = Monster.SelectSingleNode("cha")?.InnerText;
            string save = Monster.SelectSingleNode("save")?.InnerText;
            string skill = Monster.SelectSingleNode("skill")?.InnerText;
            string resist = Monster.SelectSingleNode("resist")?.InnerText;
            string immune = Monster.SelectSingleNode("immune")?.InnerText;
            string conditionimmune = Monster.SelectSingleNode("conditionImmune")?.InnerText;
            string senses = Monster.SelectSingleNode("senses")?.InnerText;
            string passive = Monster.SelectSingleNode("passive")?.InnerText;
            string languages = Monster.SelectSingleNode("languages")?.InnerText;
            string cr = ChallengeRating(Monster.SelectSingleNode("cr")?.InnerText);
            string spells = Monster.SelectSingleNode("spells")?.InnerText;
            string legendary = Monster.SelectSingleNode("legendary")?.InnerText;
            string traits = Monster.SelectSingleNode("trait")?.InnerText;
            string actions = Monster.SelectSingleNode("action")?.InnerText;
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
                XmlNodeList nodelisttrait = Monster.SelectNodes("trait");
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
                XmlNodeList nodelistactions = Monster.SelectNodes("action");
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
                XmlNodeList nodelistlegendary = Monster.SelectNodes("legendary");
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

        private void FillSpellTable(XmlNode Spell, string SpellPageID)
        {

            OneNoteApp.GetHierarchy(SpellPageID, OneNote.HierarchyScope.hsChildren, out string CurrentSpellPage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument SpellPageDocument = new XmlDocument();
            SpellPageDocument.LoadXml(CurrentSpellPage);
            XmlNamespaceManager XNSMGR = XMLMGR(SpellPageDocument);

            // single child nodes
            string name = Spell.SelectSingleNode("name")?.InnerText;
            string school = Spell.SelectSingleNode("school")?.InnerText;
            string level = Spell.SelectSingleNode("level")?.InnerText;
            string time = Spell.SelectSingleNode("time")?.InnerText;
            string range = Spell.SelectSingleNode("range")?.InnerText;
            string components = Spell.SelectSingleNode("components")?.InnerText;
            string duration = Spell.SelectSingleNode("duration")?.InnerText;
            string description = Spell.SelectSingleNode("text")?.InnerText;
            string classes = Spell.SelectSingleNode("classes")?.InnerText;
            string roll = Spell.SelectSingleNode("roll")?.InnerText;

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
                XmlNodeList DescriptionElements = Spell.SelectNodes("text");
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
            // roll
            //if (roll != null)
            //{
            //    string RollText = "";
            //    XmlNodeList RollElements = Spell.SelectNodes("roll");
            //    foreach (XmlNode rollel in RollElements)
            //    {
            //        if (rollel.InnerText != "")
            //            RollText += rollel.InnerText + Environment.NewLine;
            //    }
            //
            //    SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[roll]')]", XNSMGR).InnerText = RollText;
            //}
            //else
            //{
            //    Console.WriteLine("cleaning [roll]");
            //    XmlNode node = SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[roll]')]", XNSMGR);
            //    node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            //}

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

        private void FillRaceTable(XmlNode Race, string RacePageID)
        {
            OneNoteApp.GetHierarchy(RacePageID, OneNote.HierarchyScope.hsChildren, out string CurrentRacePage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument RacePageDocument = new XmlDocument();
            RacePageDocument.LoadXml(CurrentRacePage);
            XmlNamespaceManager XNSMGR = XMLMGR(RacePageDocument);

            // single child nodes
            string name = Race.SelectSingleNode("name")?.InnerText;
            string size = MonsterSize(Race.SelectSingleNode("size")?.InnerText.ToUpper());
            string speed = Race.SelectSingleNode("speed")?.InnerText;
            string ability = Race.SelectSingleNode("ability")?.InnerText;
            string trait = Race.SelectSingleNode("trait")?.InnerText;
            string proficiency = Race.SelectSingleNode("proficiency")?.InnerText;

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
                XmlNodeList elementlist = Race.SelectNodes("trait");
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

        private void FillFeatTable(XmlNode Feat, string FeatPageID)
        {
            OneNoteApp.GetHierarchy(FeatPageID, OneNote.HierarchyScope.hsChildren, out string CurrentFeatPage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument FeatPageDocument = new XmlDocument();
            FeatPageDocument.LoadXml(CurrentFeatPage);
            XmlNamespaceManager XNSMGR = XMLMGR(FeatPageDocument);

            // single child nodes
            string name = Feat.SelectSingleNode("name")?.InnerText;
            string modifier = Feat.SelectSingleNode("modifier")?.InnerText;
            string prerequisites = Feat.SelectSingleNode("prerequisite")?.InnerText;
            string description = Feat.SelectSingleNode("text")?.InnerText;

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

            if (description != null & description != "")
            {
                string objstring = "";
                XmlNodeList elementlist = Feat.SelectNodes("text");
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

        private void FillBackgroundTable(XmlNode InputNode, string PageID)
        {
            OneNoteApp.GetHierarchy(PageID, OneNote.HierarchyScope.hsChildren, out string CurrentPage);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml(CurrentPage);
            XmlNamespaceManager XNSMGR = XMLMGR(xmldoc);

            // single child nodes
            string name = InputNode.SelectSingleNode("name")?.InnerText;
            string trait = InputNode.SelectSingleNode("trait")?.InnerText;
            //string proficiency = InputNode.SelectSingleNode("proficiency")?.InnerText;

            // name
            if (name != null)
                xmldoc.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = boldname + name + endspan;

            if (trait != null)
            {
                string objstring = "";
                XmlNodeList elementlist = InputNode.SelectNodes("trait");
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
            if (inputCR == null | inputCR == "")
                return null;

            inputCR = inputCR.Replace("0", "0 (0 or 10 xp)");
            inputCR = inputCR.Replace("1/8", "1/8 (25 xp)");
            inputCR = inputCR.Replace("1/4", "1/4 (50 xp)");
            inputCR = inputCR.Replace("1/2", "1/2 (100 xp)");
            inputCR = inputCR.Replace("1", "1 (200 xp)");
            inputCR = inputCR.Replace("2", "2 (450 xp)");
            inputCR = inputCR.Replace("3", "3 (700 xp)");
            inputCR = inputCR.Replace("4", "4 (1,100 xp)");
            inputCR = inputCR.Replace("5", "5 (1,800 xp)");
            inputCR = inputCR.Replace("6", "6 (2,300 xp)");
            inputCR = inputCR.Replace("7", "7 (2,900 xp)");
            inputCR = inputCR.Replace("8", "8 (3,900 xp)");
            inputCR = inputCR.Replace("9", "9 (5,000 xp)");
            inputCR = inputCR.Replace("10", "10 (5,900 xp)");
            inputCR = inputCR.Replace("11", "11 (7,200 xp)");
            inputCR = inputCR.Replace("12", "12 (8,400 xp)");
            inputCR = inputCR.Replace("13", "13 (10,000 xp)");
            inputCR = inputCR.Replace("14", "14 (11,500 xp)");
            inputCR = inputCR.Replace("15", "15 (13,000 xp)");
            inputCR = inputCR.Replace("16", "16 (15,000 xp)");
            inputCR = inputCR.Replace("17", "17 (18,000 xp)");
            inputCR = inputCR.Replace("18", "18 (20,000 xp)");
            inputCR = inputCR.Replace("19", "19 (22,000 xp)");
            inputCR = inputCR.Replace("20", "20 (25,000 xp)");
            inputCR = inputCR.Replace("21", "21 (33,000 xp)");
            inputCR = inputCR.Replace("22", "22 (41,000 xp)");
            inputCR = inputCR.Replace("23", "23 (50,000 xp)");
            inputCR = inputCR.Replace("24", "24 (62,000 xp)");
            inputCR = inputCR.Replace("25", "25 (75,000 xp)");
            inputCR = inputCR.Replace("26", "26 (90,000 xp)");
            inputCR = inputCR.Replace("27", "27 (105,000 xp)");
            inputCR = inputCR.Replace("28", "28 (120,000 xp)");
            inputCR = inputCR.Replace("29", "29 (135,000 xp)");
            inputCR = inputCR.Replace("30", "30 (155,000 xp)");
            return inputCR;
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

    }
}
