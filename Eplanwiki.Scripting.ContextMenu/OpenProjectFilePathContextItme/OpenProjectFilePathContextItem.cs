/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Eplanwiki.Scripting.ContextMenu
{
    /// <summary>
    /// Copy this scriptfile to your Eplan Script folder and 
    /// load it by selecting "Load Script" in the Utilities>>Scripts menue.
    /// A new entry (in GUI language) in the page navigator contextmenu 
    /// will appear which opens the path to the selected project.
    /// </summary>
    public class OpenProjectFilePathContextItem
    {
        Eplan.EplApi.Gui.ContextMenu menu;
        ContextMenuLocation menuLocation;
        String menuText;

        /// <summary>
        /// Execution after loading  this script
        /// </summary>
        [DeclareRegister]
        public void Register()
        {
            InitiateMenu();
        }

        /// <summary>
        /// Execution after unloading this script
        /// </summary>
        [DeclareUnregister]
        public void UnRegister()
        {
            DisposeMenu();
        }

        /// <summary>
        /// Initiates the contextmenu entry in the page navigator.
        /// The method is seperated from Register() so that it is more easy to copy sections into other scriptfiles.
        /// </summary>
        public void InitiateMenu()
        {
            menuText = getMenuText();
            menu = new Eplan.EplApi.Gui.ContextMenu();
            menuLocation = new ContextMenuLocation("PmPageObjectTreeDialog", "1007");
            menu.AddMenuItem(menuLocation, menuText, "OpenProjectFilePath", false, false);
        }

        /// <summary>
        /// Removes the contextmenu entry from the page navigator.        
        /// </summary>
        public void DisposeMenu()
        {
            if (menu != null && menuLocation != null)
            {
                menu.RemoveMenuItem(menuLocation, menuText, "OpenProjectFilePath", false, false);
            }
        }

        /// <summary>
        /// Returns the menueitem text in the gui langueage if available.
        /// Translated with google :-\
        /// </summary>
        /// <returns></returns>
        private string getMenuText()
        {                        
            MultiLangString muLangMenuText = new MultiLangString();
            muLangMenuText.SetAsString(
                "de_DE@Dateipfad öffnen;" + 
                "en_US@Open Path;" +
                "ru_RU@Открыть путь;" +
                "pt_PT@Abrir Caminho;" +                
                "fr_FR@Ouvrir Chemin;"
                );
            ISOCode guiLanguage = new Languages().GuiLanguage;
            return muLangMenuText.GetString((ISOCode.Language)guiLanguage.GetNumber());
        }

        /// <summary>
        /// Action opens the projectpath in explorer
        /// </summary>
        [DeclareAction("OpenProjectFilePath")]
        public void OpenProjectFilePath()
        {
            try
            {
                Process.Start(PathMap.SubstitutePath("$(P)"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
