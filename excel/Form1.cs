using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excel
{
	public partial class Form1 : Form
	{
		private FolderBrowserDialog folderBrowser;
		public Form1()
		{
			InitializeComponent();
		}

		private void exportBtn_Click(object sender, EventArgs e)
		{
			folderBrowser = new FolderBrowserDialog();

			string fileName = "document1";
			//ask for filename
			if (InputBox.UseInputBox("New document", "New document name:", ref fileName) == DialogResult.OK)
			{
				// Browsing folder
				DialogResult result = folderBrowser.ShowDialog();
				if (result == DialogResult.OK)
				{
					string folderName = folderBrowser.SelectedPath;

					//save file
					if (Excel.SaveFile(folderName, fileName, panel1))
					{
						MessageBox.Show("Your file have been saved.", "Yay!");
						return;
					}
					else
					{
						MessageBox.Show("Excel is not properly installed!!");
						return;
					}

				}
			}		

		}

		private void createBtn_Click(object sender, EventArgs e)
		{
			panel1.Controls.Clear();
			// creating textboxes
			for (int i = 0; i < colNum.Value; i++)
			{
				for (int j = 0; j < rowNum.Value; j++)
				{
					TextBox txtBox = new TextBox();
					txtBox.Font = new Font(label1.Font.Name, 10F);
					txtBox.Size = new Size(100, 40);
					txtBox.Multiline = true;
					txtBox.Location = new Point(i * 100, j * 40);
					txtBox.BackColor = (i) % 2 == 0 ? Color.DarkSeaGreen : Color.White;
					txtBox.BorderStyle = BorderStyle.FixedSingle;
					txtBox.TextAlign = HorizontalAlignment.Center;
					txtBox.Text = null;
					txtBox.Click += txt_Click;
					txtBox.Name = string.Format("{0}_{1}", i+1, j+1);
					panel1.Controls.Add(txtBox);
					

				}
			}
		}

		private void txt_Click(object sender, EventArgs e)
		{
			TextBox txt = sender as TextBox;
			Control mainBox = SearchBox(panel1,txt.Name);
			if (mainBox == null)
				return;
			mainBox.Name = txt.Name;
		}

		public Control SearchBox(Control panel, string desiredBox)
		{
			foreach (Control box in panel.Controls)
			{
				if (box.HasChildren)
				{
					SearchBox(box, desiredBox);
				}
				else
				{
					// check if control has searched name
					if (box.Name == desiredBox)
					{
						return box;
					}
				}
			}
			return null;
		}
	}
}
