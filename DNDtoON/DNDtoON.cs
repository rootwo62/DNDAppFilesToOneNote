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
    class Load
    {

        OneNote.Application OneNoteApp = new OneNote.Application();


        string notebookID, sectionID, MonsterBlockPageID, SpellBlockPageID, pageID;
        string notebooksxmlfile, sectionsxmlfile, pagesxmlfile, monsterblocktemplatefile, spellblocktemplatefile;
        string monsterblockpagename, spellblockpagename, sectionname, notebookName;
        string bestiaryFile, spellsFile;

        static void Main(string[] args)
        {

            Load Load = new Load();

            if (Properties.Settings.Default.BlockType == "monster")
            {
                Load.GetMonsterTableXML();
                XmlDocument MonsterBestiaryXMLDoc = new XmlDocument();
                MonsterBestiaryXMLDoc.Load(Load.bestiaryFile);

                XmlNodeList Monsters = MonsterBestiaryXMLDoc.SelectNodes("//monster");

                foreach (XmlNode Monster in Monsters)
                {
                    Console.WriteLine("adding {0} to {1}", Monster.SelectSingleNode("name").InnerText, Load.sectionname);
                    Load.CopyMonterTableTemplate(Monster.SelectSingleNode("name").InnerText, out string MonsterPageID);
                    Load.FillMonsterStatsTable(Monster, MonsterPageID);
                }
            }
            if (Properties.Settings.Default.BlockType == "spell")
            {
                Load.GetSpellTableXML();
                XmlDocument SpellsXMLDoc = new XmlDocument();
                SpellsXMLDoc.Load(Load.spellsFile);

                XmlNodeList Spells = SpellsXMLDoc.SelectNodes("//spell");

                int i = 0;

                foreach (XmlNode Spell in Spells)
                {
                    i++;
                    //if (i > 30)
                    //    break;
                    Console.WriteLine("adding {0} to {1}", Spell.SelectSingleNode("name").InnerText, Load.sectionname);
                    Load.CopySpellTableTemplate(Spell.SelectSingleNode("name").InnerText, out string SpellPageID);
                    Load.FillSpellTable(Spell, SpellPageID);
                }
            }

            Console.WriteLine("press any key to close...");
            Console.ReadKey(true);
        }

        XmlNamespaceManager XMLMGR(XmlDocument InputXMLDocument)
        {
            XmlNamespaceManager XMLMGR = new XmlNamespaceManager(InputXMLDocument.NameTable);
            XMLMGR.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");
            return XMLMGR;
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

        private void GetMonsterTableXML()
        {
            bestiaryFile = Properties.Settings.Default.BestiaryFile;
            monsterblockpagename = Properties.Settings.Default.MonsterBlockTemplate;
            sectionname = Properties.Settings.Default.OneNoteSection;
            notebookName = Properties.Settings.Default.Notebook;

            notebooksxmlfile = "notebooks.xml";
            sectionsxmlfile = "sections.xml";
            pagesxmlfile = "pages.xml";
            monsterblocktemplatefile = "monsterblocktemplate.xml";

            GetONHeierarchy("", OneNote.HierarchyScope.hsNotebooks, notebooksxmlfile);
            notebookID = ID(notebooksxmlfile, notebookName);

            GetONHeierarchy(notebookID, OneNote.HierarchyScope.hsSections, sectionsxmlfile);
            sectionID = ID(sectionsxmlfile, sectionname);

            GetONHeierarchy(sectionID, OneNote.HierarchyScope.hsPages, pagesxmlfile);
            MonsterBlockPageID = ID(pagesxmlfile, monsterblockpagename);

            OneNoteApp.GetHierarchy(MonsterBlockPageID, OneNote.HierarchyScope.hsChildren, out string MonsterBlockPageXML);

            XDocument.Parse(MonsterBlockPageXML).Save(monsterblocktemplatefile);
        }

        private void GetSpellTableXML()
        {

            spellsFile = Properties.Settings.Default.SpellsFile;
            spellblockpagename = Properties.Settings.Default.SpellBlockTemplate;
            sectionname = Properties.Settings.Default.OneNoteSection;
            notebookName = Properties.Settings.Default.Notebook;

            notebooksxmlfile = "notebooks.xml";
            sectionsxmlfile = "sections.xml";
            pagesxmlfile = "pages.xml";
            spellblocktemplatefile = "spellblocktemplate.xml";

            GetONHeierarchy("", OneNote.HierarchyScope.hsNotebooks, notebooksxmlfile);
            notebookID = ID(notebooksxmlfile, notebookName);

            GetONHeierarchy(notebookID, OneNote.HierarchyScope.hsSections, sectionsxmlfile);
            sectionID = ID(sectionsxmlfile, sectionname);

            GetONHeierarchy(sectionID, OneNote.HierarchyScope.hsPages, pagesxmlfile);
            SpellBlockPageID = ID(pagesxmlfile, spellblockpagename);

            OneNoteApp.GetHierarchy(SpellBlockPageID, OneNote.HierarchyScope.hsChildren, out string SpellBlockPageXML);

            XDocument.Parse(SpellBlockPageXML).Save(spellblocktemplatefile);


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

        private void CopyMonterTableTemplate(string MonsterName, out string pageID)
        {
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            OneNoteApp.CreateNewPage(sectionID, out pageID);
            OneNoteApp.GetPageContent(pageID, out string NewPageXML);

            //XDocument.Parse(NewPageXML).Save("test2.xml");

            XmlDocument MonsterBlockDocument = new XmlDocument();
            MonsterBlockDocument.Load(monsterblocktemplatefile);
            XmlNode TableElement = MonsterBlockDocument.SelectSingleNode("//one:Table", XMLMGR(MonsterBlockDocument));

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

            xmldoc.SelectSingleNode("//one:Title//one:T", XMLMGR(xmldoc)).InnerText = MonsterName;



            //PageXML.AppendChild(importNode);

            //xmldoc.SelectSingleNode("//one:Outline/@objectID", XMLMGR(xmldoc)).InnerXml = "{" + Guid.NewGuid().ToString().ToUpper() + "}";

            //xmldoc.Save("test.xml");
            //Console.WriteLine(xmldoc.OuterXml);
            //OneNoteApp.UpdateHierarchy(xmldoc.OuterXml);
            OneNoteApp.UpdatePageContent(xmldoc.OuterXml);
            OneNoteApp.SyncHierarchy(sectionID);

        }

        private void CopySpellTableTemplate(string SpellName, out string pageID)
        {
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            OneNoteApp.CreateNewPage(sectionID, out pageID);
            OneNoteApp.GetPageContent(pageID, out string NewPageXML);

            //XDocument.Parse(NewPageXML).Save("test2.xml");

            XmlDocument SpellBlockDocument = new XmlDocument();
            SpellBlockDocument.Load(spellblocktemplatefile);
            XmlNode TableElement = SpellBlockDocument.SelectSingleNode("//one:Table", XMLMGR(SpellBlockDocument));

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

            xmldoc.SelectSingleNode("//one:Title//one:T", XMLMGR(xmldoc)).InnerText = SpellName;



            //PageXML.AppendChild(importNode);

            //xmldoc.SelectSingleNode("//one:Outline/@objectID", XMLMGR(xmldoc)).InnerXml = "{" + Guid.NewGuid().ToString().ToUpper() + "}";

            //xmldoc.Save("test.xml");
            //Console.WriteLine(xmldoc.OuterXml);
            //OneNoteApp.UpdateHierarchy(xmldoc.OuterXml);
            OneNoteApp.UpdatePageContent(xmldoc.OuterXml);
            OneNoteApp.SyncHierarchy(sectionID);

        }

        private string MonsterSize(string inputSize)
        {
            inputSize = inputSize.Replace("S", "Small");
            inputSize = inputSize.Replace("M", "Medium");
            inputSize = inputSize.Replace("L", "Large");
            inputSize = inputSize.Replace("H", "Huge");
            inputSize = inputSize.Replace("G", "Gargantuan");
            return inputSize;
        }

        private XmlDocument Blah(XmlDocument inputDocument, XmlNodeList nodeList, XmlNode onenoteNodeToUpdate)
        {
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            var font = @"<span style = 'font-size:10pt; font-family:cambria;'>";
            var bold = @"<span style = 'font-weight:bold; font-size:10pt; font-family:Cambria'>";
            var italic = @"<span style='font-style:italic'>";
            var endspan = @"</span>";

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
                traits += bold + statusTrait.SelectSingleNode("name").InnerText + endspan + font + attack + endspan + Environment.NewLine;

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
            string cr = Monster.SelectSingleNode("cr")?.InnerText;
            string spells = Monster.SelectSingleNode("spells")?.InnerText;
            string legendary = Monster.SelectSingleNode("legendary")?.InnerText;
            string traits = Monster.SelectSingleNode("trait")?.InnerText;
            string actions = Monster.SelectSingleNode("action")?.InnerText;
            string enviornment = "any";

            size = MonsterSize(size);

            var font = @"<span style = 'font-size:9pt' 'font-family:cambria'>";
            var bold = @"<span style = 'font-weight:bold' 'font-size:9pt'>";
            var italic = @"<span style='font-style:italic'>";
            var endspan = @"</span>";

            // name
            if (name != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = bold + name + endspan;
            // size
            if (size != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[size]')]", XNSMGR).InnerText = italic + size + ", " + type + ", " + alignment + endspan;
            // armor class
            if (ac != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[ac]')]", XNSMGR).InnerText = bold + @"Armor Class: " + endspan + ac;
            // hit points
            if (hp != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[hp]')]", XNSMGR).InnerText = bold + @"Hit Points: " + endspan + hp;
            // speed
            if (speed != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[speed]')]", XNSMGR).InnerText = bold + @"Speed: " + endspan + speed;
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
            if (save != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[saving throws]')]", XNSMGR).InnerText = bold + @"Saving Throws: " + endspan + save;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[saving throws]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // skills
            if (skill != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[skill]')]", XNSMGR).InnerText = skill;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[skill]')]", XNSMGR);
                node.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.RemoveChild(node.ParentNode.ParentNode.ParentNode.ParentNode);
            }
            // resist
            if (resist != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[resist]')]", XNSMGR).InnerText = bold + @"Resistance: " + endspan + resist;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[resist]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // immune
            if (immune != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[immune]')]", XNSMGR).InnerText = bold + @"Damage Immunities: " + endspan + immune;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[immune]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // condition immunities
            if (conditionimmune != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[condition immunities]')]", XNSMGR).InnerText = bold + @"Condition Immunities: " + endspan + conditionimmune;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[condition immunities]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // passive
            if (passive != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[passive]')]", XNSMGR).InnerText = bold + @"Passive Perception: " + endspan + passive;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[passive]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // senses
            if (senses != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[senses]')]", XNSMGR).InnerText = bold + @"Senses: " + endspan + senses;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[senses]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // languages
            if (languages != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[languages]')]", XNSMGR).InnerText = bold + @"Languages: " + endspan + languages;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[languages]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // challenge rating
            if (cr != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[cr]')]", XNSMGR).InnerText = bold + @"Challenge: " + endspan + cr;
            else
            {
                XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[cr]')]", XNSMGR);
                node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
            }
            // enviornment
            if (enviornment != null)
                MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[environment]')]", XNSMGR).InnerText = enviornment;
            // spells
            if (spells != null)
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
                MonsterPageXMLDoc = Blah(MonsterPageXMLDoc, nodelisttrait, nodetraits);
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
                MonsterPageXMLDoc = Blah(MonsterPageXMLDoc, nodelistactions, nodeactions);
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
            if (Monster.SelectSingleNode("legendary") != null)
            {
                XmlNodeList nodelistlegendary = Monster.SelectNodes("legendary");
                XmlNode nodelegendary = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[legendary]')]", XNSMGR).ParentNode.ParentNode;
                MonsterPageXMLDoc = Blah(MonsterPageXMLDoc, nodelistlegendary, nodelegendary);
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

            var font = @"<span style = 'font-size:9pt' 'font-family:cambria'>";
            var bold = @"<span style = 'font-weight:bold' 'font-size:9pt'>";
            var italic = @"<span style='font-style:italic'>";
            var endspan = @"</span>";

            // name
            if (name != null)
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[name]')]", XNSMGR).InnerText = bold + name + endspan;
            // school
            if (school != null)
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[school]')]", XNSMGR).InnerText = italic + level + ", " + school + endspan;
            // level
            //if (level != null)
            //    SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[level]')]", XNSMGR).InnerText = level;
            // time
            if (time != null)
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[time]')]", XNSMGR).InnerText = bold + "Time to Cast: " + endspan + time;
            // range
            if (range != null)
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[range]')]", XNSMGR).InnerText = bold + "Spell Range: " + endspan + range;
            // components
            if (components != null)
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[components]')]", XNSMGR).InnerText = bold + "Components: " + endspan + components;
            // duration
            if (duration != null)
                SpellPageDocument.SelectSingleNode("//one:T[contains(text(), '[duration]')]", XNSMGR).InnerText = bold + "Cast Time: " + endspan + duration;
            // description
            if (description != null)
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
            if (classes != null)
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

        private bool ContainsChildren(XmlNodeList Nodes)
        {
            if (Nodes.Count > 1)
                return true;

            return false;
        }

    }
}
