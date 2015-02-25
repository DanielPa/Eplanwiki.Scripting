using Eplan.EplApi.Scripting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using Eplan.EplApi.Base;
using System.Xml.XPath;
using Eplan.EplApi.ApplicationFramework;
using System.Drawing;

namespace Eplanwiki.Scripting.ContextMenu
{
    public class SetDipSwitchAddress
    {
        [DeclareAction("SetDipSwitchAddress", 20)]
        public void Run()
        {
            string address = string.Empty;
            string macroPath = PathMap.SubstitutePath("$(MD_MACROS)");
            string macroFilePath = macroPath + "\\DIP.ema";
            string tempPath = PathMap.SubstitutePath("$(TMP)");
            string tempMacro = tempPath + "\\tempDIP.ema";
            string textToReplace = "??_??@{{ADDRESS}};";
            int bitsNeeded = 14;

            if (SetDipSwitchAddress.InputBox("Insert adderss", "", ref address) == DialogResult.OK)
            {                
                int DecAddress = Convert.ToInt32(address);                
                string BinAddress = GetBinaryAdress(bitsNeeded, DecAddress);
                string DipAddress = BinAddress.Replace("1", "▀ ").Replace("0", "▄ ");

                File.Copy(macroFilePath, tempMacro, true);

                ReplaceXmlAttributeValue(tempMacro, "O30", "A511", textToReplace, "??_??@{{" + DipAddress + "}};");
                InsertMacro(tempMacro);
            }
        }

        private string GetBinaryAdress(int addressRangeInBit, int address)
        {
            string binaryString = Convert.ToString(address, 2).PadLeft(addressRangeInBit, '0'); 
            
            return binaryString;
        }
                
        private void ReplaceXmlAttributeValue(string xmlFileName, 
                                                string nodeName, 
                                                string attributeName, 
                                                string oldValue, 
                                                string newValue) 
        {         
            XmlDocument document = new XmlDocument();
            document.Load(xmlFileName);
            XmlNodeList nodeList = document.SelectNodes("//"+nodeName);

            foreach (XmlNode node in nodeList)
            {
                if (node.Attributes[attributeName].Value == oldValue)//
                {
                    node.Attributes[attributeName].Value = newValue;
                }
            }

            document.PreserveWhitespace = true;
            XmlTextWriter writer = new XmlTextWriter(xmlFileName, Encoding.UTF8);
            document.WriteTo(writer);
            writer.Close();                                   
        }
        
        private void InsertMacro(string macroFilePath)
        {
            ActionCallingContext XGedStartInteractionActionContext = new ActionCallingContext();
            XGedStartInteractionActionContext.AddParameter("Name", "XMIaInsertMacro");
            XGedStartInteractionActionContext.AddParameter("filename", macroFilePath);
            XGedStartInteractionActionContext.AddParameter("variant", "0");
            XGedStartInteractionActionContext.AddParameter("RepresentationType", "1");
            new CommandLineInterpreter().Execute("XGedStartInteractionAction", XGedStartInteractionActionContext);
        }

        private static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();            
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;
            
            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }
    }
}
