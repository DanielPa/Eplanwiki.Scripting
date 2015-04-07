using System;
using System.Xml;

namespace Eplanwiki.Scripting.EditMacroboxes
{
    /// <summary>
    /// Represents a macrobox with all properties needed in this tool
    /// </summary>
    public class MacroBox
    {
        #region Properties
        /// <summary>
        /// Read only property initialized by constructor
        /// </summary>
        public string Name { get; private set; }
        /// <summary>
        /// Property initialized by project or page
        /// </summary>
        public string NewName { get; set; }        
        /// <summary>
        /// Read only property initialized by constructor
        /// </summary>
        public string Variant { get; private set; }
        string _NewVariant;
        /// <summary>
        /// Property initialized by project or page
        /// </summary>
        public string NewVariant
        {
            get
            {
                return _NewVariant;
            }
            set
            {
                if (Convert.ToInt16(value) > 15)
                {
                    //TODO: Error message or somthing (variant has to be 0-15)                    
                    this._NewVariant = "15";
                }
                else
                {
                    this._NewVariant = value;
                }
            }
        }
        /// <summary>
        /// Read only property initialized by constructor
        /// </summary>
        public string RepresentationType { get; private set; }
        /// <summary>
        /// Property initialized by project or page
        /// </summary>
        public string NewRepresentationType { get; set; }
        /// <summary>
        /// Read only property initialized by constructor
        /// </summary>
        public string Description { get; private set; }
        /// <summary>
        /// Read only property initialized by constructor
        /// </summary>
        public string PageName { get; private set; }
        /// <summary>
        /// Read only property initialized by constructor
        /// </summary>
        public string X { get; private set; }
        /// <summary>
        /// Read only property initialized by constructor
        /// </summary>
        public string Y { get; private set; }
        #endregion

        #region Constructors
        /// <summary>
        /// Costructor used to read the needed values from 
        /// the O37 element to locate the macrobox position 
        /// in the project and compare the current values with 
        /// thoes to set.
        /// </summary>
        /// <param name="o37"></param>
        public MacroBox(XmlElement o37, string _pageName)
        {
            //TODO: maybe extract methodes from regions
            #region Set Name
            try
            {
                this.Name = o37.SelectSingleNode("P11/@P23001").Value;
            }
            catch (NullReferenceException)
            {
                this.Name = "";
            }
            #endregion

            #region Set Variant
            this.Variant = o37.GetAttribute("A690");
            #endregion

            #region Set DocumentType
            this.RepresentationType = o37.GetAttribute("A689");
            #endregion

            #region Set Description
            try
            {                
                this.Description = EplanScriptHelper.GetMuLanStringToDisplay(o37.SelectSingleNode("P11/@P23004").Value);
            }
            catch (NullReferenceException)
            {
                this.Description = "";
            }
            #endregion

            #region Set X and Y
            try
            {
                string location = o37.SelectSingleNode("S40x1201/@A762").Value;
                this.X = location.Split('/')[0];
                this.Y = location.Split('/')[1];
            }
            catch (NullReferenceException)
            {
                this.X = "";
                this.Y = "";
            }
            #endregion

            this.PageName = _pageName;
        }

        /// <summary>
        /// Default constructor, should never be used!
        /// </summary>
        public MacroBox()
        {
            throw new System.NotImplementedException();
        }
        #endregion

        #region Methods
        public void Select()
        {
            if(EplanScriptHelper.Edit(this.PageName, this.X, this.Y))
            {
                EplanScriptHelper.XGedSelect();
            }
        }

        public void SetNewName()
        {
            EplanScriptHelper.XEsSetPropertyAction("23001", "0", this.NewName);
        }

        public void SetNewVariant()
        {
            EplanScriptHelper.XEsSetPropertyAction("23008", "0", this.NewVariant);
        }

        public void SetNewRepresentationType()
        {
            EplanScriptHelper.XEsSetPropertyAction("23007", "0", this.NewRepresentationType);
        }        
        #endregion

        #region Enums
        /// <summary>
        /// Equal to Eplan::EplApi::DataModel::MasterData::WindowMacro::Enums::RepresentationType Enumeration
        /// </summary>
        public enum RepresentationTypes
        {
            Neutral = 0,
            MultiLine = 1,
            SingleLine = 2,
            PairCrossReference = 3,
            Overview = 4,
            Graphics = 5,
            ArticlePlacement = 6,
            PI_FlowChart = 7,
            Fluid_MultiLine = 8,
            Cabling = 9,
            ArticlePlacement3D = 10,
            Functional = 11,
            Planning = 12,
            Default = -1
        }
        #endregion
    }
}
