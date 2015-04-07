using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Eplanwiki.Scripting.EditMacroboxes
{
    public class EplanScriptHelper
    {
        private const ISOCode.Language LANGUAGE = ISOCode.Language.L_de_DE;        

        public void RenameMacroBoxesAfterPageStructure()
        {            
            string epjFile = ExportCurrentProject();
            Project prj;
            if (File.Exists(epjFile))
            {
                prj = new Project(epjFile);
            }
            else
            {
                MessageBox.Show("File not found!", epjFile, MessageBoxButtons.OK, MessageBoxIcon.Error );
                return;
            }
            //TODO: Show changes in dialog            
            foreach (Page page in prj.PageList)
            {
                foreach (MacroBox box in page.MacroBoxList)
                {                   
                    box.Select();
                    box.SetNewName();
                    box.SetNewVariant();
                    box.SetNewRepresentationType();
                }
                page.Close();
            }
            //TODO: Export and create protocol
        }

        private string ExportCurrentProject()
        {
            //TODO: Show progress
            string projectPath = PathMap.SubstitutePath("$(P)");
            string exportFile = projectPath + "\\XML\\Export.epj";
            ActionCallingContext exportContext = new ActionCallingContext();
            exportContext.AddParameter("TYPE", "PXFPROJECT");
            exportContext.AddParameter("PROJECTNAME", projectPath);
            exportContext.AddParameter("EXPORTFILE", exportFile);
            new CommandLineInterpreter().Execute("export", exportContext);
            return exportFile;
        }

        public static bool XEsSetPropertyAction(string propId, string index, string value)
        {            
            ActionCallingContext XEsSetPropertyActionContext = new ActionCallingContext();
            XEsSetPropertyActionContext.AddParameter("PropertyId", propId);
            XEsSetPropertyActionContext.AddParameter("PropertyIndex", index);
            XEsSetPropertyActionContext.AddParameter("PropertyValue", value);
            return new CommandLineInterpreter().Execute("XEsSetPropertyAction", XEsSetPropertyActionContext);
        }

        public static bool Edit(string page, string x, string y)
        {
            ActionCallingContext editContext = new ActionCallingContext();
            editContext.AddParameter("PAGENAME", page);
            editContext.AddParameter("X", x);
            editContext.AddParameter("Y", y);
            return new CommandLineInterpreter().Execute("edit", editContext);
        }

        public static void XGedSelect()
        {
            string windowname = Process.GetCurrentProcess().MainWindowTitle;
            //TODO: Determine Electric P8 version and set class name
            IntPtr eplanHandle = FindWindow("AfxMDIFrame110u", windowname);

            // Verify that Calculator is a running process.
            if (eplanHandle == IntPtr.Zero)
            {
                MessageBox.Show(windowname);
                return;
            }
            // Make Calculator the foreground application and send it 
            // a set of calculations.
            SetForegroundWindow(eplanHandle);
            SendKeys.SendWait(" ");
        }

        public static void XGedClosePage()
        {
            new CommandLineInterpreter().Execute("XGedClosePage");
        }

        public static string GetMuLanStringToDisplay(string mulangString)
        {
            MultiLangString mls = new MultiLangString();
            mls.SetAsString(mulangString);
            return mls.GetStringToDisplay(LANGUAGE);
        }
        
        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string lpClassName,
            string lpWindowName);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
