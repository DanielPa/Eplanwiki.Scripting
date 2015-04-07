using System;
using System.Collections.Generic;
using System.Xml;

namespace Eplanwiki.Scripting.EditMacroboxes
{
    /// <summary>
    /// Represents a page with all properties needed in this tool
    /// </summary>
    public class Page
    {
        #region Properties        
        /// <summary>
        /// Read only properties initialized by constructor
        /// </summary>
        public Dictionary<StructureSegments, string> PageStructureSegments { get; private set; }
        public string Description { get; private set; }
        public string DocumentType { get; private set; }
        public List<MacroBox> MacroBoxList { get; private set; }
        public string FullPageName { get; private set; }
        #endregion

        #region Fields
        private List<StructureSegments> _StructureSegmentOrder;

        /// <summary>
        /// Reading the values from xml is to complex!!!
        /// See xml path /EplanPxfRoot/O14/P11/@P10018
        /// </summary>
        private const string PREFIX_DESIGNATION_PLANT = "=";
        private const string PREFIX_DESIGNATION_LOCATION = "+";
        private const string PREFIX_DESIGNATION_FUNCTIONALASSIGNMENT = "==";
        private const string PREFIX_DESIGNATION_PLACEOFINSTALLATION = "++";
        private const string PREFIX_DESIGNATION_DOCTYPE = "&";
        private const string PREFIX_DESIGNATION_USERDEFINED = "#";
        private const string PREFIX_DESIGNATION_INSTALLATIONNUMBER = "_";
        private const string PREFIX_PAGE_COUNTER = "/";
        #endregion

        #region Constructors
        /// <summary>
        /// Costructor used to read the needed values from 
        /// the O4 element for setting page dependent properties 
        /// of the macroboxes located on a page.
        /// </summary>
        public Page(XmlTextReader reader, List<StructureSegments> StructureSegmentOrder)
        {
            XmlDocument document = new XmlDocument();
            document.LoadXml(reader.ReadOuterXml());
             
            //TODO: maybe extract methodes from regions
            #region Set PageStructureSements
            this.PageStructureSegments = new Dictionary<StructureSegments, string>();
            XmlNode p49 = document.SelectSingleNode("O4/P49");
            foreach (XmlAttribute attr in p49.Attributes)
            {
                int id = Convert.ToInt32(attr.Name.Remove(0,1));
                PageStructureSegments.Add((StructureSegments)id, attr.Value);                
            }
            #endregion

            #region Set Description
            try 
	        {                
                this.Description = EplanScriptHelper.GetMuLanStringToDisplay(document.SelectSingleNode("O4/P11/@P11011").Value);
	        }
	        catch (NullReferenceException)
	        {
                this.Description = "";
            }
            #endregion
            
            #region Set DocumentType
            this.DocumentType = ((XmlElement)document.SelectSingleNode("O4")).GetAttribute("A47");
            #endregion

            //How is this solved in entity frameworks? 
            _StructureSegmentOrder = StructureSegmentOrder;

            SetFullPageName();

            #region Set MacroBoxList
            this.MacroBoxList = new List<MacroBox>();
            foreach (XmlNode o37 in document.SelectNodes("O4/O37"))
            {
                this.MacroBoxList.Add(new MacroBox((XmlElement)o37, this.FullPageName));
            }
            #endregion            

            SetMacroBoxNewNames();            
            SetMacroBoxNewRepresentationTypes();            
        }

        /// <summary>
        /// Default constructor, should never be used!
        /// </summary>
        public Page()
        {
            throw new System.NotImplementedException();
        }
        #endregion

        #region Methods
        /// <summary>
        /// Close page in eplan GED
        /// </summary>
        public void Close()
        {
            EplanScriptHelper.XGedClosePage();
        }
       
        /// <summary>
        /// Full pagename needed for edit-action parameter
        /// </summary>
        private void SetFullPageName()
        {
            foreach (StructureSegments segment in _StructureSegmentOrder)
            {
                if (this.PageStructureSegments.ContainsKey(segment))
                {
                    switch (segment)
                    {
                        case StructureSegments.DESIGNATION_PLANT:
                            this.FullPageName += PREFIX_DESIGNATION_PLANT + this.PageStructureSegments[segment];
                            break;
                        case StructureSegments.DESIGNATION_LOCATION:
                            this.FullPageName += PREFIX_DESIGNATION_LOCATION + this.PageStructureSegments[segment];
                            break;
                        case StructureSegments.DESIGNATION_FUNCTIONALASSIGNMENT:
                            this.FullPageName += PREFIX_DESIGNATION_FUNCTIONALASSIGNMENT + this.PageStructureSegments[segment];
                            break;
                        case StructureSegments.DESIGNATION_PLACEOFINSTALLATION:
                            this.FullPageName += PREFIX_DESIGNATION_PLACEOFINSTALLATION + this.PageStructureSegments[segment];
                            break;
                        case StructureSegments.DESIGNATION_DOCTYPE:
                            this.FullPageName += PREFIX_DESIGNATION_DOCTYPE + this.PageStructureSegments[segment];
                            break;
                        case StructureSegments.DESIGNATION_USERDEFINED:
                            this.FullPageName += PREFIX_DESIGNATION_USERDEFINED + this.PageStructureSegments[segment];
                            break;
                        case StructureSegments.DESIGNATION_INSTALLATIONNUMBER:
                            this.FullPageName += PREFIX_DESIGNATION_INSTALLATIONNUMBER + this.PageStructureSegments[segment];
                            break;                        
                    }
                }
            }
            this.FullPageName += PREFIX_PAGE_COUNTER + this.PageStructureSegments[StructureSegments.PAGE_COUNTER];
        }

        /// <summary>
        /// Set macro path by page structure and name by page description
        /// </summary>
        private void SetMacroBoxNewNames()
        {
            foreach (MacroBox box in this.MacroBoxList)
            {
                string newName = "";
                foreach (StructureSegments segment in _StructureSegmentOrder)
                {
                    if (this.PageStructureSegments.ContainsKey(segment))
                    {
                        newName += this.PageStructureSegments[segment] + "\\";
                    }                    
                }
                newName += this.Description;
                box.NewName = newName;
            }
        }
              
        /// <summary>
        /// Set the repType of a macrobox equal to the located page (if possible, else MultiLine)
        /// </summary>
        private void SetMacroBoxNewRepresentationTypes()
        {
            MacroBox.RepresentationTypes setToType;

            try
            {
                switch ((DocumentTypes)Convert.ToInt32(this.DocumentType))
                {
                    case DocumentTypes.PanelLayout: setToType = MacroBox.RepresentationTypes.ArticlePlacement;
                        break;
                    case DocumentTypes.PanelLayout3D: setToType = MacroBox.RepresentationTypes.ArticlePlacement3D;
                        break;
                    case DocumentTypes.CableConnectionDiagram: setToType = MacroBox.RepresentationTypes.Cabling;
                        break;
                    /*case null: setToType = MacroBox.RepresentationTypes.Default;
                        break;*/
                    case DocumentTypes.CircuitFluid: setToType = MacroBox.RepresentationTypes.Fluid_MultiLine;
                        break;
                    case DocumentTypes.Functional: setToType = MacroBox.RepresentationTypes.Functional;
                        break;
                    case DocumentTypes.Graphics: setToType = MacroBox.RepresentationTypes.Graphics;
                        break;
                    case DocumentTypes.Circuit: setToType = MacroBox.RepresentationTypes.MultiLine;
                        break;
                    /*case null: setToType = MacroBox.RepresentationTypes.Neutral;
                        break;*/
                    case DocumentTypes.Overview: setToType = MacroBox.RepresentationTypes.Overview;
                        break;
                    /*case null: setToType = MacroBox.RepresentationTypes.PairCrossReference;
                        break;*/
                    /*case null: setToType = MacroBox.RepresentationTypes.PI_FlowChart;
                        break;*/
                    case DocumentTypes.Planning: setToType = MacroBox.RepresentationTypes.Planning;
                        break;
                    case DocumentTypes.CircuitSingleLine: setToType = MacroBox.RepresentationTypes.SingleLine;
                        break;
                    default: setToType = MacroBox.RepresentationTypes.MultiLine;
                        break;
                }

                foreach (MacroBox box in this.MacroBoxList)
                {
                    box.NewRepresentationType = Convert.ToInt16(setToType).ToString();                    
                }
            }
            catch (Exception)
            {
                foreach (MacroBox box in this.MacroBoxList)
                {
                    box.NewRepresentationType = box.RepresentationType;
                }
            }
        }
        #endregion

        #region Enums        
        /// <summary>
        /// Structure segments a page can have, as propID values
        /// </summary>
        public enum StructureSegments
	    {
	        DESIGNATION_PLANT = 1100,                   
            DESIGNATION_LOCATION = 1200,                
            DESIGNATION_FUNCTIONALASSIGNMENT = 1300,    
            DESIGNATION_PLACEOFINSTALLATION = 1400,     
            DESIGNATION_DOCTYPE = 1500,                 
            DESIGNATION_USERDEFINED = 1600,             
            DESIGNATION_INSTALLATIONNUMBER = 1700,      
            PAGE_COUNTER = 11012                        
	    }

        /// <summary>
        /// Equal to Eplan::EplApi::DataModel::DocumentTypeManager::DocumentType Enumeration
        /// </summary>
        public enum DocumentTypes 
        {
            Undefined = 0,
            Circuit = 1,
            CircuitSingleLine = 2,
            Overview = 3,
            CableLayout = 4,
            ExternalDocument = 5,
            LogocadTriga = 6,
            Graphics = 7,
            PanelLayout = 8,
            TerminalDiagram = 9,
            TerminalConnectiondiagram = 10,
            InterconnectDiagram = 11,
            TerminalLineupDiagram = 12,
            DeviceList = 13,
            TableOfContents = 14,
            TerminalStripOverview = 15,
            PLCCardOverview = 16,
            TitlePage = 17,
            SymbolOverview = 18,
            ConnectionList = 19,
            PotentialOverview = 20,
            CableOverview = 21,
            PartsList = 22,
            PlugOverview = 23,
            PLCDiagram = 24,
            DeviceConnectionDiagram = 25,
            CableConnectionDiagram = 26,
            PlugConnectionDiagram = 27,
            PlugDiagram = 28,
            PanelLayoutCaption = 29,
            PartsSumList = 30,
            StructIdentifierOverview = 31,
            FormOverview = 32,
            FrameOverview = 33,
            CircuitFluid = 34,
            RevisionOverview = 35,
            ProcessAndInstrumentationDiagram = 38,
            ModelView = 40,
            Topology = 43,
            Planning = 54,
            Functional = -12,
            PanelLayout3D = -8,
            PanelLayoutDetail = -6,
            Symbol = -5,
            Form = -4,
            External = -3,
            PairCrossReference = -2,
            Frame = -1
        }
        #endregion        
    }
}
