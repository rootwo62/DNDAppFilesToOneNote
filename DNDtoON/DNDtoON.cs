using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Threading;
using OneNote = Microsoft.Office.Interop.OneNote;
using System.Text.RegularExpressions;
using System.Globalization;
using System.IO;

namespace DNDtoON
{
    class Application
    {

        OneNote.Application OneNoteApp = new OneNote.Application();

        string notebookID, rootID, sectionID, pageID;
        string rootxmlfile, sectionsxmlfile, pagesxmlfile;
        string sectionname, notebookName, rootName;

        bool debug = Properties.Settings.Default.Debug;

        const string font = @"<span style = 'font-size:9pt; font-family:cambria'>";
        const string boldfont = @"<span style = 'font-weight:bold; font-size:9t; font-family:cambria'>";
        const string boldname = @"<span style = 'font-weight:bold; font-size:14pt; font-family:cambria'>";
        const string tableheaderblack = @"<span style = 'color:#FFFFFF; font-weight:bold; font-size:9pt; font-family:cambria;'>";
        const string boldtraitname = @"<span style = 'color:#000000; font-weight:bold; font-size:9pt; font-family:cambria;'>";
        const string tablefont = @"<span style = 'color:#000000; font-size:8pt; font-family:cambria;'>";
        const string italicfont = @"<span style='font-style:italic; font-size:9pt; font-family:cambria'>";
        const string bolditalicfont = @"<span style='font-style:italic; font-weight:bold; font-size:9pt; font-family:cambria'>";
        const string endspan = @"</span>";

        const string ConsoleTitle = "DNDtoON";

        static void Main(string[] args)
        {
            Console.Title = ConsoleTitle;
            Application Application = new Application();
            Application.Run();

            //Application.Run();

            Console.WriteLine("press any key to close...");
            Console.ReadKey(true);
        }

        private void test()
        {
            OneNoteApp.GetHierarchy(OneNoteApp.Windows.CurrentWindow.CurrentSectionGroupId, OneNote.HierarchyScope.hsChildren, out string currentsection);
            XDocument.Parse(currentsection).Save("debug-currentview.xml");
        }

        private void Run()
        {
            string Compendium = Properties.Settings.Default.DNDAppFileXML;

            XDocument xdocCompendium = XDocument.Load(Compendium);

            string blocktype = Properties.Settings.Default.BlockType;
            string templatexmlfile = blocktype + "templatefile.xml";
            GetOneNoteTableXML(Properties.Settings.Default.ONBlockTemplatePageName, templatexmlfile);

            Console.WriteLine("ROOT: {1}{0}INITIAL SECTION: {2}{0}COMPENDIUM: {3}{0}Press any key to continue...{0}", Environment.NewLine, Properties.Settings.Default.ONRootPath, sectionname, Compendium);
            Console.ReadKey(true);

            int xdocCompendiumCount = xdocCompendium.Descendants(blocktype).Count();
            int currentElementIndex = 0;

            foreach (XElement element in xdocCompendium.Descendants(blocktype))
            {

                //Console.WriteLine("source book = {0}", SourceBook(element.Descendants("type").First().Value));

                currentElementIndex++;
                string newtitle = string.Format("{0} | {1}/{2}", ConsoleTitle, currentElementIndex, xdocCompendiumCount);
                Console.Title = newtitle;


                if (Properties.Settings.Default.CopyPageToSourceBookSection)
                    SetSection(SourceBook(element.Descendants("type").First().Value, sectionname), rootID);
                else
                    SetSection(sectionname, rootID);

                string pageName = element.DescendantsAndSelf("name").First().Value;
                CopyPageTableTemplate(pageName, templatexmlfile, out string pageID);

                Console.WriteLine("adding {0} to section {1} in notebook {2}", pageName, sectionname, rootName);

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

                if (Properties.Settings.Default.SlowProcess | debug)
                    Console.ReadKey(true);
            }

            if (File.Exists(rootxmlfile))
                File.Delete(rootxmlfile);
            if (File.Exists(sectionsxmlfile))
                File.Delete(sectionsxmlfile);
            if (File.Exists(pagesxmlfile))
                File.Delete(pagesxmlfile);
            if (File.Exists(templatexmlfile))
                File.Delete(templatexmlfile);
        }

        private void GetONHierarchyFile(string ID, OneNote.HierarchyScope OneNoteHierarchyScope, string OutputFile)
        {
            OneNoteApp.GetHierarchy(ID, OneNoteHierarchyScope, out string pbstrHierarchyXMLOut);
            XDocument.Parse(pbstrHierarchyXMLOut).Save(OutputFile);
        }

        // Load Variables For Specified Block

        private void GetOneNoteTableXML(string templatePageName, string outXMLFile)
        {
            rootxmlfile = "root.xml";
            sectionsxmlfile = "sections.xml";
            pagesxmlfile = "pages.xml";
            notebookName = Properties.Settings.Default.ONRootPath;
            sectionname = Properties.Settings.Default.ONSection;

            // check the notebook name for a sectiongroup and the assign the onenote root
            if (notebookName.Contains("/"))
            {
                notebookName = Properties.Settings.Default.ONRootPath.Split('/')[0];
                rootName = Properties.Settings.Default.ONRootPath.Split('/')[1];

                OneNoteApp.GetHierarchy("", OneNote.HierarchyScope.hsNotebooks, out string Notebooks);
                notebookID = (from el in XDocument.Parse(Notebooks).Root.Elements()
                              where el.Attribute("name").Value == notebookName
                              select el).FirstOrDefault().Attribute("ID").Value;

                OneNoteApp.GetHierarchy(notebookID, OneNote.HierarchyScope.hsChildren, out string NotebookElements);
                XDocument.Parse(NotebookElements).Save(rootxmlfile);
                rootID = (from el in XDocument.Parse(NotebookElements).Root.Elements()
                          where el.Attribute("name").Value == rootName
                          select el).FirstOrDefault().Attribute("ID").Value;
            }
            else
            {
                rootName = notebookName;
                OneNoteApp.GetHierarchy("", OneNote.HierarchyScope.hsNotebooks, out string Notebooks);
                notebookID = (from el in XDocument.Parse(Notebooks).Root.Elements()
                              where el.Attribute("name").Value == notebookName
                              select el).FirstOrDefault().Attribute("ID").Value;

                OneNoteApp.GetHierarchy(notebookID, OneNote.HierarchyScope.hsChildren, out string NotebookElements);
                XDocument.Parse(Notebooks).Save(rootxmlfile);
                rootID = notebookID;
            }

            try
            {
                // get the pages from the selected root and select the desired template
                OneNoteApp.GetHierarchy(rootID, OneNote.HierarchyScope.hsPages, out string RootPages);
                XDocument.Parse(RootPages).Save(pagesxmlfile);
                pageID = (from el in XDocument.Parse(RootPages).Root.DescendantsAndSelf().Elements()
                          where el.Attribute("name").Value == templatePageName
                          select el).FirstOrDefault().Attribute("ID").Value;

                if (debug)
                    Console.WriteLine("notebook name = '{1}', notebook id = {2}{0}" +
                                    "root name = '{3}', root id = {4}{0}" +
                                    "page name = '{5}', page id = {6}{0}", Environment.NewLine, notebookName, notebookID, rootName, rootID, templatePageName, pageID);

                //Console.ReadKey(true);
                OneNoteApp.GetPageContent(pageID, out string templateXML, OneNote.PageInfo.piAll);
                XDocument.Parse(templateXML).Save(outXMLFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("the template {0} was not found in {1}", templatePageName, rootName);
                Console.ReadKey(true);
                if (debug)
                    Console.WriteLine(ex.Message);
                throw;
            }

        }

        // Copy Templates

        private void CopyPageTableTemplate(string outPageName, string templatePageXML, out string newPageID)
        {
            XNamespace URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";
            OneNoteApp.CreateNewPage(sectionID, out newPageID);
            OneNoteApp.GetPageContent(newPageID, out string NewPageXML);

            XDocument xdocFromFileTemplate = XDocument.Load(templatePageXML);
            string TemplatPageID = xdocFromFileTemplate.Root.Attribute("ID").Value;
            OneNoteApp.GetPageContent(TemplatPageID, out string rawtemplatexml, OneNote.PageInfo.piAll);
            XDocument xdocTemplate = XDocument.Parse(rawtemplatexml);

            if (debug)
                XDocument.Parse(rawtemplatexml).Save("debug-rawtemplate.xml");

            XDocument docNewPage = null;
            docNewPage = XDocument.Parse(NewPageXML);
            docNewPage.Root.Element(URI + "PageSettings").Attribute("color").Value = "#FEF8E6";

            // copy background from template if it exsists
            if (xdocTemplate.Root.Element(URI + "Image") != null)
            {
                docNewPage.Root.Add(xdocTemplate.Root.Element(URI + "Image"));
                docNewPage.Root.Element(URI + "Image").Attribute("objectID").Remove();
            }

            foreach (XElement Outline in xdocTemplate.Root.Elements(URI + "Outline"))
            {
                //i++;
                //Console.WriteLine("outline{0}", i);
                XElement xelTemplateOutline = Outline;
                xelTemplateOutline.Attribute("objectID").Remove();
                docNewPage.Root.Add(xelTemplateOutline);
                docNewPage.Root.Descendants(URI + "Title").First().Descendants(URI + "T").First().Value = outPageName;
            }

            if (docNewPage.Root.Attribute("stationeryName") == null)
                docNewPage.Root.Add(new XAttribute("stationeryName", "DNDPAGE"));

            docNewPage.Root.Attribute("name").Value = outPageName;

            try
            {
                if (debug)
                    docNewPage.Save(string.Format("{0}.xml", "debug-newpage.xml"));

                OneNoteApp.UpdatePageContent(docNewPage.ToString());
                OneNoteApp.SyncHierarchy(sectionID);
            }
            catch (Exception ex)
            {
                Console.WriteLine("failed to copy the template for {0}", outPageName);
                if (debug)
                    Console.WriteLine(ex.Message.Trim());
            }

        }

        // Fill Tables

        private void FillMonsterStatsTable(string InputXML, string MonsterPageID)
        {

            OneNoteApp.GetHierarchy(MonsterPageID, OneNote.HierarchyScope.hsChildren, out string CurrentMonsterPageXml);
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlDocument Monster = new XmlDocument();
            Monster.LoadXml(InputXML);

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
            string strength = Modifier(Monster.SelectSingleNode("//str")?.InnerText);
            string dexterity = Modifier(Monster.SelectSingleNode("//dex")?.InnerText);
            string constitution = Modifier(Monster.SelectSingleNode("//con")?.InnerText);
            string intelligence = Modifier(Monster.SelectSingleNode("//int")?.InnerText);
            string wisdom = Modifier(Monster.SelectSingleNode("//wis")?.InnerText);
            string charisma = Modifier(Monster.SelectSingleNode("//cha")?.InnerText);
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
            string link = PageTitleLink(MonsterPageID, name);

            //Console.ReadKey(true);
            size = MonsterSize(size);

            try
            {
                // initiative outline
                if (MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), 'Initiative')]", XNSMGR) != null)
                {
                    if (link != null)
                        MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[ct-name]')]", XNSMGR).InnerText = link;
                    if (ac != null)
                        MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[ct-ac]')]", XNSMGR).InnerText = ac.Split('(')[0].Trim();
                    if (hp != null)
                        MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[ct-hpcurrent]')]", XNSMGR).InnerText = hp.Split('(')[0].Trim();
                    if (hp != null)
                        MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[ct-hpmax]')]", XNSMGR).InnerText = hp.Split('(')[0].Trim();
                    if (cr != null & cr != "")
                        MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[ct-cr]')]", XNSMGR).InnerText = cr.Split('(')[1].TrimEnd(')').Trim();
                    else
                    {
                        XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[ct-cr]')]", XNSMGR);
                        node.InnerText = "";
                    }
                }

                // outline 1

                // name
                if (name != null) { }
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
                    MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[skills]')]", XNSMGR).InnerText = boldfont + @"Skills: " + endspan + skill;
                else
                {
                    XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[skills]')]", XNSMGR);
                    node.ParentNode.ParentNode.RemoveChild(node.ParentNode);
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
                    MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[spells]')]", XNSMGR).InnerText = boldfont + @"Spells: " + endspan + spells;
                else
                {
                    XmlNode node = MonsterPageXMLDoc.SelectSingleNode("//one:T[contains(text(), '[spells]')]", XNSMGR);
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
                }
                MonsterPageXMLDoc.SelectSingleNode("//one:Outline/one:Size/@width", XNSMGR).InnerXml = "400";

                XDocument.Parse(MonsterPageXMLDoc.OuterXml).Save("output-monsterpagexmldoc.xml");

                OneNoteApp.UpdatePageContent(MonsterPageXMLDoc.OuterXml);
            }
            catch (Exception ex)
            {
                if (debug)
                {
                    if (!Directory.Exists("errors"))
                        Directory.CreateDirectory("errors");
                    XDocument.Parse(MonsterPageXMLDoc.OuterXml).Save("errors/" + name + "-onenote.xml");
                    XDocument.Parse(Monster.OuterXml).Save("errors/" + name + "-monster.xml");
                }
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
                if (debug)
                {
                    if (!Directory.Exists("errors"))
                        Directory.CreateDirectory("errors");
                    XDocument.Parse(SpellPageDocument.OuterXml).Save("errors/" + name + "-onenote.xml");
                    XDocument.Parse(Spell.OuterXml).Save("errors/" + name + "-spells.xml");
                }
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

        private string GetIDByName(string XMLFILE, string inputAttribute)
        {
            //XNamespace URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";
            try
            {
                return (from el in XDocument.Load(XMLFILE).Root.Elements()
                        where el.Attribute("name").Value == inputAttribute
                        select el).FirstOrDefault().Attribute("ID").Value;
            }
            catch (Exception)
            {
                return "";
            }
        }

        private string PageTitleLink(string PageID, string PageName)
        {
            XNamespace URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";
            OneNoteApp.GetHyperlinkToObject(PageID, "", out string hyperlink);
            string link = string.Format("<a href='{0}'>{1}</a>", hyperlink, PageName);
            return link;
        }

        private string Modifier(string inputAttribute)
        {
            int x = Convert.ToInt32(inputAttribute);

            //Console.WriteLine(x);

            if (x == 1)
                return x + " (-5)";
            if (x >= 2 && x <= 3)
                return x + " (-4)";
            if (x >= 4 && x <= 5)
                return x + " (-3)";
            if (x >= 6 && x <= 7)
                return x + " (-2)";
            if (x >= 8 && x <= 9)
                return x + " (-1)";
            if (x >= 10 && x <= 11)
                return x + " (0)";
            if (x >= 12 && x <= 13)
                return x + " (1)";
            if (x >= 14 && x <= 15)
                return x + " (2)";
            if (x >= 16 && x <= 17)
                return x + " (3)";
            if (x >= 18 && x <= 19)
                return x + " (4)";
            if (x >= 20 && x <= 21)
                return x + " (5)";
            if (x >= 22 && x <= 23)
                return x + " (6)";
            if (x >= 24 && x <= 25)
                return x + " (7)";
            if (x >= 26 && x <= 27)
                return x + " (8)";
            if (x >= 28 && x <= 29)
                return x + " (9)";

            return "";

        }

        private XmlDocument MonsterTableCleanup(XmlDocument inputDocument, XmlNodeList nodeList, XmlNode onenoteNodeToUpdate)
        {
            string URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            XmlNode oneOE = inputDocument.CreateElement("one:OE", URI);
            XmlNode oneT = inputDocument.CreateElement("one:T", URI);
            oneT.InnerText = "";
            oneOE.AppendChild(oneT);
            onenoteNodeToUpdate.AppendChild(oneOE);

            foreach (XmlNode statusTrait in nodeList)
            {
                string traits = "";
                string attack = "";


                oneOE = inputDocument.CreateElement("one:OE", URI);
                oneT = inputDocument.CreateElement("one:T", URI);
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

                // add the trait name
                traits += Environment.NewLine + boldtraitname + statusTrait.SelectSingleNode("name").InnerText + endspan + font + attack + endspan + Environment.NewLine;

                // get list of <text> elements form monster bestiery
                XmlNodeList nodelistText = statusTrait.SelectNodes("text");
                foreach (XmlNode nodetraittext in nodelistText)
                {
                    if (nodetraittext.InnerText != "")
                    {
                        string nodetext = nodetraittext.InnerText;

                        if (nodetext.Contains("•"))
                            nodetext = nodetext.Replace("•", "    • ");

                        traits += font + nodetext + endspan + Environment.NewLine;
                    }
                   
                }

                //traits += Environment.NewLine;
                oneT.InnerText = traits;
                oneOE.AppendChild(oneT);
                onenoteNodeToUpdate.AppendChild(oneOE);

                oneOE = inputDocument.CreateElement("one:OE", URI);
                oneT = inputDocument.CreateElement("one:T", URI);
                oneT.InnerText = "";
                oneOE.AppendChild(oneT);
                onenoteNodeToUpdate.AppendChild(oneOE);


            }

            return inputDocument;
        }

        private string MonsterSize(string inputSize)
        {
            if (inputSize == null | inputSize == "")
                return null;

            if (inputSize == "T")
                return "Tiny";
            if (inputSize == "S")
                return "Small";
            if (inputSize == "M")
                return "Medium";
            if (inputSize == "L")
                return "Large";
            if (inputSize == "H")
                return "Huge";
            if (inputSize == "G")
                return "Gargantuan";

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

        private string SourceBook(string inputDescription, string otherNameIfMissing)
        {
            if (inputDescription == "" | inputDescription == null)
                return otherNameIfMissing;

            string typedescription = inputDescription;
            string checkparentheses =  @"(\(.*\))";
            string regexparentheses = Regex.Match(inputDescription, checkparentheses).Value;
            if (regexparentheses.Contains(","))
                typedescription = typedescription.Replace(regexparentheses, "");

            TextInfo output = new CultureInfo("en-US", false).TextInfo;

            return output.ToTitleCase(typedescription.Split(',')[1].Trim());
        }

        private void SetSection(string ONSectionName, string ONRootID)
        {
            XNamespace URI = "http://schemas.microsoft.com/office/onenote/2013/onenote";
            OneNoteApp.GetHierarchy(rootID, OneNote.HierarchyScope.hsSections, out string sectionsxml);
            XDocument xdoc = XDocument.Parse(sectionsxml);
            xdoc.Save(sectionsxmlfile);

            try
            {
                sectionID = (from el in XDocument.Parse(sectionsxml).Root.DescendantsAndSelf().Elements(URI + "Section")
                             where el.Attribute("name").Value == ONSectionName
                             select el).FirstOrDefault().Attribute("ID").Value;
            }
            catch (Exception)
            {

                sectionID = "";
            }


            if (sectionID.Length <= 0)
            {
                xdoc.Root.Add(new XElement(URI + "Section", new XAttribute("name", ONSectionName)));
                xdoc.Save(sectionsxmlfile);

                OneNoteApp.UpdateHierarchy(xdoc.ToString());
                OneNoteApp.SyncHierarchy(rootID);

                OneNoteApp.GetHierarchy(rootID, OneNote.HierarchyScope.hsSections, out string newrootdescendentxml);

                XDocument.Parse(newrootdescendentxml).Save(sectionsxmlfile);
                sectionID = GetIDByName(sectionsxmlfile, ONSectionName);

                Console.WriteLine("created new section {0}", ONSectionName);

                if (debug)
                {
                    OneNoteApp.GetHierarchy(sectionID, OneNote.HierarchyScope.hsChildren, out string newsectionxml);
                    XDocument.Parse(newsectionxml).Save("debug-newsection.xml");
                }

                //Console.ReadKey(true);
            }
            
            
            //Console.WriteLine("sectionname {0}'s ID = {1}", SectionName, sectionID);
            
        }

    }
}
