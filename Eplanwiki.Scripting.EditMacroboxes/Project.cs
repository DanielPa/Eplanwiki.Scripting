using System;
using System.Collections.Generic;
using System.Xml;
using System.Linq;

namespace Eplanwiki.Scripting.EditMacroboxes
{
    /// <summary>
    /// The only class which needs to be created after generating xml from eplan project.
    /// The rest is itteration over lists!
    /// </summary>
    public class Project
    {
        #region Properties
        public List<Page> PageList { get; set; }
        public List<Page.StructureSegments> StructureSegmentOrder { get; set; }
        public List<MacroBox> AllMacroBoxesList { get; set; }
        #endregion

        #region Constructors
        /// <summary>
        /// Construcor used to get info from xml (epj-export)
        /// </summary>
        /// <param name="epjPath"></param>
        public Project(string epjPath)
        {
            XmlTextReader reader = new XmlTextReader(epjPath);
            reader.WhitespaceHandling = WhitespaceHandling.None;
            this.PageList = new List<Page>();            

            //TODO: maybe extract methodes from regions
            #region Set StructureSegmentOrder

            //Realy cool and fast way to get info from xml with combining XmlTextReader and DOM
            //see also: http://www.codeproject.com/Articles/156982/How-to-Open-Large-XML-files-without-Loading-the-XM
            while (reader.Read())
            {
                               
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "O14" && reader.IsStartElement())
                {
                    this.StructureSegmentOrder = new List<Page.StructureSegments>();
                    reader.Read();
                    if (reader.LocalName == "P11" && reader.HasAttributes)
                    {                        
                        for (int i = 1; i < 10; i++)
                        {
                            string attrName = "P100" + i.ToString("D2");
                            if (reader.GetAttribute(attrName) != null)
                            {
                                Int32 str = Convert.ToInt32(reader.GetAttribute(attrName).Split('>')[0].Remove(0, 1));
                                this.StructureSegmentOrder.Add((Page.StructureSegments) str);
                            }
                        }                 
                    }
                }
                                              
            }
            reader.Close();
                #endregion

            #region Set PageList
            reader = new XmlTextReader(epjPath);
            while (reader.Read())
            {                
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "O4" && reader.IsStartElement())
                {                    
                    this.PageList.Add(new Page(reader, this.StructureSegmentOrder));
                }                
            }
            reader.Close();
            #endregion

            SetAllMacroBoxVariants();
        }

        /// <summary>
        /// Default constructor, should never be used!
        /// </summary>
        public Project()
        {
            throw new System.NotImplementedException();
        }
        #endregion

        #region Methods
        /// <summary>
        /// Has to be page superior because variants could be distibuted over several pages.
        /// Variant will raise if same NewName with same repType already exists in list. 
        /// </summary>
        private void SetAllMacroBoxVariants()
        {
            this.AllMacroBoxesList = new List<MacroBox>();
            foreach (Page page in PageList)
            {
                foreach (MacroBox box in page.MacroBoxList)
                {
                    box.NewVariant = AllMacroBoxesList.Count(o => o.NewName == box.NewName && o.RepresentationType == box.RepresentationType).ToString();
                    AllMacroBoxesList.Add(box);
                }
            }
        }
        #endregion
    }
}
