using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using System.IO;
using Tekla.Technology.Akit.UserScript;
using Tekla.Technology.Scripting;
using System.Diagnostics;
using dr = Tekla.Structures.Drawing;
using sr = System.Reflection;
using Microsoft.Win32;


namespace Tekla.Technology.Akit.UserScript
{
  partial class FindAssForm: System.Windows.Forms.Form
    {

        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.выбратьВМоделиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.показатьВыбранноеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.скрытьВыбранноеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.обнвитьПространствоМоделиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.OpenChert = new System.Windows.Forms.ToolStripMenuItem();
            this.labelMO = new System.Windows.Forms.LinkLabel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageAssPos = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.AssButton = new System.Windows.Forms.Button();
            this.FindSW = new System.Windows.Forms.ComboBox();
            this.FindText = new System.Windows.Forms.ComboBox();
            this.FindButton = new System.Windows.Forms.Button();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.tabPageIDGuid = new System.Windows.Forms.TabPage();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.linkLabelMOh = new System.Windows.Forms.LinkLabel();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.textBoxID = new System.Windows.Forms.TextBox();
            this.contextMenuStrip1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPageAssPos.SuspendLayout();
            this.tabPageIDGuid.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.выбратьВМоделиToolStripMenuItem,
            this.показатьВыбранноеToolStripMenuItem,
            this.скрытьВыбранноеToolStripMenuItem,
            this.обнвитьПространствоМоделиToolStripMenuItem,
            this.toolStripSeparator1,
            this.OpenChert});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(274, 142);
            // 
            // выбратьВМоделиToolStripMenuItem
            // 
            this.выбратьВМоделиToolStripMenuItem.Name = "выбратьВМоделиToolStripMenuItem";
            this.выбратьВМоделиToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.Insert;
            this.выбратьВМоделиToolStripMenuItem.Size = new System.Drawing.Size(273, 22);
            this.выбратьВМоделиToolStripMenuItem.Text = "Select in model";
            this.выбратьВМоделиToolStripMenuItem.Click += new System.EventHandler(this.выбратьВМоделиToolStripMenuItem_Click);
            // 
            // показатьВыбранноеToolStripMenuItem
            // 
            this.показатьВыбранноеToolStripMenuItem.Name = "показатьВыбранноеToolStripMenuItem";
            this.показатьВыбранноеToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.H)));
            this.показатьВыбранноеToolStripMenuItem.Size = new System.Drawing.Size(273, 22);
            this.показатьВыбранноеToolStripMenuItem.Text = "Show only selected";
            this.показатьВыбранноеToolStripMenuItem.Click += new System.EventHandler(this.показатьВыбранноеToolStripMenuItem_Click);
            // 
            // скрытьВыбранноеToolStripMenuItem
            // 
            this.скрытьВыбранноеToolStripMenuItem.Name = "скрытьВыбранноеToolStripMenuItem";
            this.скрытьВыбранноеToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.U)));
            this.скрытьВыбранноеToolStripMenuItem.Size = new System.Drawing.Size(273, 22);
            this.скрытьВыбранноеToolStripMenuItem.Text = "Hide selected";
            this.скрытьВыбранноеToolStripMenuItem.Click += new System.EventHandler(this.скрытьВыбранноеToolStripMenuItem_Click);
            // 
            // обнвитьПространствоМоделиToolStripMenuItem
            // 
            this.обнвитьПространствоМоделиToolStripMenuItem.Name = "обнвитьПространствоМоделиToolStripMenuItem";
            this.обнвитьПространствоМоделиToolStripMenuItem.Size = new System.Drawing.Size(273, 22);
            this.обнвитьПространствоМоделиToolStripMenuItem.Text = "Reload model space";
            this.обнвитьПространствоМоделиToolStripMenuItem.Click += new System.EventHandler(this.обнвитьПространствоМоделиToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(270, 6);
            // 
            // OpenChert
            // 
            this.OpenChert.Name = "OpenChert";
            this.OpenChert.Size = new System.Drawing.Size(273, 22);
            this.OpenChert.Text = "Open Drawing";
            this.OpenChert.Click += new System.EventHandler(this.OpenChert_Click);
            // 
            // labelMO
            // 
            this.labelMO.AutoSize = true;
            this.labelMO.Location = new System.Drawing.Point(17, 52);
            this.labelMO.Name = "labelMO";
            this.labelMO.Size = new System.Drawing.Size(0, 13);
            this.labelMO.TabIndex = 7;
            this.labelMO.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.labelMO_LinkClicked);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageAssPos);
            this.tabControl1.Controls.Add(this.tabPageIDGuid);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(337, 287);
            this.tabControl1.TabIndex = 10;
            // 
            // tabPageAssPos
            // 
            this.tabPageAssPos.Controls.Add(this.button1);
            this.tabPageAssPos.Controls.Add(this.progressBar1);
            this.tabPageAssPos.Controls.Add(this.AssButton);
            this.tabPageAssPos.Controls.Add(this.FindSW);
            this.tabPageAssPos.Controls.Add(this.FindText);
            this.tabPageAssPos.Controls.Add(this.FindButton);
            this.tabPageAssPos.Controls.Add(this.treeView1);
            this.tabPageAssPos.Location = new System.Drawing.Point(4, 22);
            this.tabPageAssPos.Name = "tabPageAssPos";
            this.tabPageAssPos.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageAssPos.Size = new System.Drawing.Size(329, 261);
            this.tabPageAssPos.TabIndex = 0;
            this.tabPageAssPos.Text = "Search parts/assembly";
            this.tabPageAssPos.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(310, 243);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(18, 19);
            this.button1.TabIndex = 10;
            this.button1.Text = "x";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.CancelOperation);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.progressBar1.Location = new System.Drawing.Point(3, 243);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(307, 18);
            this.progressBar1.TabIndex = 9;
            // 
            // AssButton
            // 
            this.AssButton.Location = new System.Drawing.Point(3, 10);
            this.AssButton.Name = "AssButton";
            this.AssButton.Size = new System.Drawing.Size(118, 24);
            this.AssButton.TabIndex = 8;
            this.AssButton.Tag = "True";
            this.AssButton.Text = "Search Assembly";
            this.AssButton.UseVisualStyleBackColor = true;
            this.AssButton.Click += new System.EventHandler(this.AssButton_Click_1);
            // 
            // FindSW
            // 
            this.FindSW.FormattingEnabled = true;
            this.FindSW.Location = new System.Drawing.Point(3, 39);
            this.FindSW.Name = "FindSW";
            this.FindSW.Size = new System.Drawing.Size(119, 21);
            this.FindSW.TabIndex = 7;
            // 
            // FindText
            // 
            this.FindText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FindText.Location = new System.Drawing.Point(128, 40);
            this.FindText.Name = "FindText";
            this.FindText.Size = new System.Drawing.Size(195, 21);
            this.FindText.TabIndex = 4;
            this.FindText.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.FindText_KeyPress_1);
            // 
            // FindButton
            // 
            this.FindButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FindButton.Location = new System.Drawing.Point(128, 10);
            this.FindButton.Name = "FindButton";
            this.FindButton.Size = new System.Drawing.Size(195, 24);
            this.FindButton.TabIndex = 5;
            this.FindButton.Text = "Search";
            this.FindButton.UseVisualStyleBackColor = true;
            this.FindButton.Click += new System.EventHandler(this.FindButF_Click);
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeView1.ContextMenuStrip = this.contextMenuStrip1;
            this.treeView1.FullRowSelect = true;
            this.treeView1.Location = new System.Drawing.Point(3, 66);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(323, 175);
            this.treeView1.TabIndex = 6;
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // tabPageIDGuid
            // 
            this.tabPageIDGuid.Controls.Add(this.linkLabel2);
            this.tabPageIDGuid.Controls.Add(this.linkLabelMOh);
            this.tabPageIDGuid.Controls.Add(this.labelMO);
            this.tabPageIDGuid.Controls.Add(this.linkLabel1);
            this.tabPageIDGuid.Controls.Add(this.textBoxID);
            this.tabPageIDGuid.Location = new System.Drawing.Point(4, 22);
            this.tabPageIDGuid.Name = "tabPageIDGuid";
            this.tabPageIDGuid.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageIDGuid.Size = new System.Drawing.Size(329, 161);
            this.tabPageIDGuid.TabIndex = 1;
            this.tabPageIDGuid.Text = "Search by ID/GUID";
            this.tabPageIDGuid.UseVisualStyleBackColor = true;
            // 
            // linkLabel2
            // 
            this.linkLabel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(246, 145);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(80, 13);
            this.linkLabel2.TabIndex = 12;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // linkLabelMOh
            // 
            this.linkLabelMOh.AutoSize = true;
            this.linkLabelMOh.LinkColor = System.Drawing.Color.Silver;
            this.linkLabelMOh.Location = new System.Drawing.Point(17, 78);
            this.linkLabelMOh.Name = "linkLabelMOh";
            this.linkLabelMOh.Size = new System.Drawing.Size(0, 13);
            this.linkLabelMOh.TabIndex = 11;
            this.linkLabelMOh.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelMOh_LinkClicked);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.LinkColor = System.Drawing.Color.Red;
            this.linkLabel1.Location = new System.Drawing.Point(17, 13);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(50, 13);
            this.linkLabel1.TabIndex = 10;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "ID/GUID";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // textBoxID
            // 
            this.textBoxID.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxID.Location = new System.Drawing.Point(73, 10);
            this.textBoxID.Name = "textBoxID";
            this.textBoxID.Size = new System.Drawing.Size(227, 20);
            this.textBoxID.TabIndex = 9;
            this.textBoxID.TextChanged += new System.EventHandler(this.textBoxID_TextChanged);
            // 
            // FindAssForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(337, 287);
            this.Controls.Add(this.tabControl1);
            this.Name = "FindAssForm";
            this.Text = "Search Assembly";
            this.TopMost = true;
            this.contextMenuStrip1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPageAssPos.ResumeLayout(false);
            this.tabPageIDGuid.ResumeLayout(false);
            this.tabPageIDGuid.PerformLayout();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem выбратьВМоделиToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem показатьВыбранноеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem скрытьВыбранноеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem обнвитьПространствоМоделиToolStripMenuItem;
        private System.Windows.Forms.LinkLabel labelMO;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageAssPos;
        private System.Windows.Forms.ComboBox FindSW;
        private System.Windows.Forms.ComboBox FindText;
        private System.Windows.Forms.Button FindButton;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.TabPage tabPageIDGuid;
        private System.Windows.Forms.LinkLabel linkLabelMOh;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.TextBox textBoxID;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem OpenChert;
        private System.Windows.Forms.Button AssButton;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button button1;
	
		Model M = new Model();
        String sysPath = "";
        System.ComponentModel.BackgroundWorker backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
           
        public FindAssForm()
        {
            InitializeComponent();
            sysPath = M.GetInfo().ModelPath;
            sysPath = Path.Combine(sysPath, "attributes");
            sysPath = Path.Combine(sysPath, "AF.SObjGrp");
            FindSW.Items.Add("Begins with");
            FindSW.Items.Add("Contains");
            FindSW.Items.Add("Equal");
            FindSW.Items.Add("Not contains");
            FindSW.Items.Add("Ends with");
            FindSW.Items.Add("Not equal");
            FindSW.Text = "Begins with";
            AssButton.Tag = true;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
            LoadItems();
        }

        public void CancelOperation(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
            progressBar1.Value=0;
        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ModelObjectEnumerator ME = M.GetModelObjectSelector().GetObjectsByFilterName("AF");
            ME.SelectInstances = false;
            string ASSEMBLY_POSITION = "";
            string ASSEMBLY_NAME = "";
            string DR_name = "";
            int sch = 0;
            if (ME.GetSize()==0)
            { return; }
            Action<int> act_max = set_max;
            Invoke(act_max, ME.GetSize());
            Action<TreeNodeCollection, TreeNode, string> act_add = Add_Node;
            Action<TreeNodeCollection, TreeNode, string, int> act_ins = Ins_Node;
            while (ME.MoveNext())
            {
                ++sch;
                backgroundWorker1.ReportProgress(sch);
                if ((sender as BackgroundWorker).CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }                              
                ModelObject tAss = ME.Current;               
                if (tAss is Part & !(bool)AssButton.Tag | tAss is Assembly & (bool)AssButton.Tag)
                {
                    if (tAss is Part & !(bool)AssButton.Tag)
                    {
                        tAss.GetReportProperty("PART_POS", ref ASSEMBLY_POSITION);
                        tAss.GetReportProperty("ASSEMBLY_POS", ref ASSEMBLY_NAME);
                        tAss.GetReportProperty("DRAWING.NAME", ref DR_name);
                        ASSEMBLY_POSITION = ASSEMBLY_POSITION.PadRight(10, ' ') + ASSEMBLY_NAME.PadRight(10, ' ') + DR_name;
                    }
                    if (tAss is Assembly & (bool)AssButton.Tag)
                    {
                        tAss.GetReportProperty("ASSEMBLY_POS", ref ASSEMBLY_POSITION);
                        tAss.GetReportProperty("DRAWING.NAME", ref DR_name);
                        ASSEMBLY_POSITION = ASSEMBLY_POSITION.PadRight(10, ' ') + DR_name;
                    }
                    TreeNodeCollection favNode = treeView1.Nodes;
                    int insIndex = treeView1.Nodes.Count;
                    TreeNode[] fNode = treeView1.Nodes.Find(ASSEMBLY_POSITION, false);
                    if (fNode.Count() > 0)
                    {
                        TreeNode grNode = fNode[0];
                        favNode = grNode.Nodes;
                        if (grNode.Tag is Part|grNode.Tag is Assembly)
                        {
                            TreeNode ngNode = new TreeNode();
                            ngNode.Tag = grNode.Tag;
                            grNode.Tag = "GR" + (ngNode.Tag as ModelObject).Identifier.ID.ToString();
                            grNode.Name = grNode.Text;
                            Invoke(act_add, new object[] { grNode.Nodes, ngNode, grNode.Text });
                        }
                    }
                    else
                    {
                        foreach (TreeNode tn in treeView1.Nodes)
                        {
                            if (GetSortString(tn.Text).CompareTo(GetSortString(ASSEMBLY_POSITION)) > 0)
                            {
                                insIndex = treeView1.Nodes.IndexOf(tn);
                                break;
                            }
                        }
                    }
                    TreeNode nNode = new TreeNode();
                    nNode.Tag = tAss;
                    nNode.Name = ASSEMBLY_POSITION;
                    Invoke(act_ins, new object[] { favNode, nNode, ASSEMBLY_POSITION, insIndex });  
                }
            }            
        }
        private string GetSortString(string s)
        {
            string sOut = s;

            string Num = "";
            string Str = "";
            int addZ = 4;
            if (sOut.IndexOf(" - ") > -1)
            {
                addZ = 4 - (sOut.Length - sOut.IndexOf(" - ") - 3);
                sOut = sOut.Replace(" - ", "");
            }

            foreach (char Nm in sOut)
            {
                if (!Char.IsNumber(Nm))
                {
                    Str = Str + Nm;
                }
                else
                {
                    Num = Num + Nm;
                }
            }
            sOut = Str + Num.PadLeft(20 + (4 - addZ), '0') + "".PadRight(addZ, '0');
            return sOut;
        }

        private void set_max(int max)
        {
            progressBar1.Maximum = max;
        }
        private void Add_Node(TreeNodeCollection fn,TreeNode cn,string nameG)
        {            
            cn.Text = nameG;
            fn.Add(cn);            
        }
        private void Ins_Node(TreeNodeCollection fn, TreeNode cn, string nameG,int idx)
        {
            cn.Text = nameG;
            fn.Insert(idx, cn);
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = 0;
            Cursor = Cursors.Default;
        }

        private void FindButF_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            CreateFilterFile();
            treeView1.Nodes.Clear();           
            backgroundWorker1.RunWorkerAsync();
            if ((bool)AssButton.Tag) { RunMacros("select_assemblies"); }
            else { RunMacros("sel_objects_in_joints"); }
            if (FindText.Items.IndexOf(FindText.Text)==-1)
            {
                FindText.Items.Add(FindText.Text);
                SaveItems();
            }
        }

        private void SaveItems()
        {
            StreamWriter sr = new StreamWriter(Path.Combine(M.GetInfo().ModelPath, "Attributes") + @"\findAss.txt",false,Encoding.Default);           
            foreach (object keyValuePair in FindText.Items)
            {
                sr.WriteLine(keyValuePair);
            }
            sr.Close();            
        }
        private void LoadItems()
        {
            try
            {
                string[] lines = File.ReadAllLines(Path.Combine(M.GetInfo().ModelPath, "Attributes") + @"\findAss.txt", Encoding.Default);
                FindText.Items.AddRange(lines);
            }
            catch
            {

            }
                
        }
        private void CreateFilterFile()
        {
          
            StreamWriter FF = new StreamWriter(sysPath,false,Encoding.Default);
            FF.WriteLine("TITLE_OBJECT_GROUP");
            FF.WriteLine("{");
            FF.WriteLine("    Version= 1.05");
            FF.WriteLine("    Count= 2 ");
            FF.WriteLine("    SECTION_OBJECT_GROUP");
            FF.WriteLine("    {");
            FF.WriteLine("        0");
            FF.WriteLine("        1");
            if ((bool)AssButton.Tag)
            {
                FF.WriteLine("        co_assembly");
                FF.WriteLine("        proASSEMBLY_POSITION");
                FF.WriteLine("        albl_Position_number");
            }
            else
            {
                FF.WriteLine("        co_part");
                FF.WriteLine("        proNUMBERING_POSITION");
                FF.WriteLine("        albl_Position_number");
            }
            switch (FindSW.Text)
            {
                case "Равно":
                    FF.WriteLine("        ==");
                    FF.WriteLine("        albl_Equals");
                    break;
                case "Начинается с":
                    FF.WriteLine("        *BEGINS*");
                    FF.WriteLine("        albl_BeginsWith");
                    break;
                case "Содержит":
                    FF.WriteLine("        *CONTAINS*");
                    FF.WriteLine("        albl_Contains");
                    break;
                case "Не содержит":
                    FF.WriteLine("        *NOTCONTAINS*");
                    FF.WriteLine("        albl_DoesNotContain");
                    break;
                case "Заканчивается на":
                    FF.WriteLine("        *ENDS*");
                    FF.WriteLine("        albl_EndsWith");
                    break;
                case "Не равно":
                    FF.WriteLine("        !=");
                    FF.WriteLine("        albl_DoesNotEqual");
                    break;
                default:
                    FF.WriteLine("        ==");
                    FF.WriteLine("        albl_Equals");
                    break;
            }
            FF.WriteLine("        " + FindText.Text.Replace((Char)32, (Char)2));
            FF.WriteLine("        0");
            FF.WriteLine("        &&");
            FF.WriteLine("        }");
            FF.WriteLine(" SECTION_OBJECT_GROUP ");
            FF.WriteLine("    {");
            FF.WriteLine("        0 ");
            FF.WriteLine("       1 ");
            FF.WriteLine("       co_object ");
            FF.WriteLine("        proOBJECT_TYPE ");
            FF.WriteLine("       albl_ObjectType ");
            FF.WriteLine("        == ");
            FF.WriteLine("        albl_Equals ");
            if ((bool)AssButton.Tag)
            {
                FF.WriteLine("        albl_Assembly ");
            }
            else
            {
                FF.WriteLine("        albl_Part ");
            }
            FF.WriteLine("        0 ");
            FF.WriteLine("        && ");
            FF.WriteLine("        }");            
            FF.WriteLine("    }");
            FF.Close();
        }

        private void SelectNode(TreeNode tn)
        {
            if (tn == null)
            {
                return;
            }
            if (tn.Tag is Assembly | tn.Tag is Part | tn.Tag.ToString().Substring(0,2) == "GR")
            {
                ArrayList sel = new ArrayList();
                if (tn.Tag is Assembly|tn.Tag is Part)
                {
                    sel.Add(tn.Tag);
                }
                else
                {
                    foreach (TreeNode trn in tn.Nodes)
                    {
                        sel.Add(trn.Tag);
                    }
                }
                Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                MS.Select(sel);               
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            SelectNode(e.Node);
            Tekla.Structures.ModelInternal.Operation.dotStartAction("ZoomToSelected", "");
        }

        private void обнвитьПространствоМоделиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Tekla.Structures.ModelInternal.Operation.dotStartAction("RedrawAll", "");
        }

        private void выбратьВМоделиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = treeView1.SelectedNode;
            SelectNode(tn);
            Tekla.Structures.ModelInternal.Operation.dotStartAction("ZoomToSelected", "");
        }

        private void показатьВыбранноеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = treeView1.SelectedNode;
            SelectNode(tn);
            Tekla.Structures.ModelInternal.Operation.dotStartAction("HideUnselectedObjects", "");
        }

        private void скрытьВыбранноеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = treeView1.SelectedNode;
            SelectNode(tn);
            Tekla.Structures.ModelInternal.Operation.dotStartAction("HideSelectedObjects", "");
        }

        private void FindText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar==(Char)13)
            {
                FindButF_Click(null, null);
            }
        }

        private void textBoxID_TextChanged(object sender, EventArgs e)
        {
            Int32 intID = 0;
            if (textBoxID.Text.Length == 38)
            {
                textBoxID.Text = textBoxID.Text.Remove(0, 2);
            }
            if (textBoxID.Text.Length == 36)
            {
                intID = 0;
                try
                {
                    Identifier id = M.GetIdentifierByGUID(textBoxID.Text);
                    intID = id.ID;
                    if (intID==0)
                    {
                        labelMO.Text = "";
                        labelMO.Tag = null;
                        linkLabelMOh.Text = "";
                        linkLabelMOh.Tag = null;
                        return;
                    }
                }
                catch
                {
                    labelMO.Text = "";
                    labelMO.Tag = null;
                    linkLabelMOh.Text = "";
                    linkLabelMOh.Tag = null;
                    return;
                }
            }

            if (textBoxID.Text.Length < 36)
            {
                try
                {
                    intID = Int32.Parse(textBoxID.Text);
                }
                catch
                {
                    labelMO.Text = "";
                    labelMO.Tag = null;
                    linkLabelMOh.Text = "";
                    linkLabelMOh.Tag = null;
                    return;
                }
            }

            Identifier ID = new Identifier(intID);
            ModelObject mo = M.SelectModelObject(ID);
            if (mo != null)
            {
                labelMO.Text = "Search in model " + mo.ToString();
                labelMO.Tag = mo;
                linkLabelMOh.Text = "Hide model except " + mo.ToString();
                linkLabelMOh.Tag = mo;
            }
            else
            {
                labelMO.Text = "";
                labelMO.Tag = null;
                linkLabelMOh.Text = "";
                linkLabelMOh.Tag = null;
            }
        }

        private void labelMO_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (labelMO.Tag is ModelObject)
            {
                ArrayList sel = new ArrayList();
                sel.Add(labelMO.Tag);
                Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                MS.Select(sel);
                Tekla.Structures.ModelInternal.Operation.dotStartAction("ZoomToSelected", "");
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (labelMO.Tag is ModelObject)
            {
                if (MessageBox.Show("Удалить объект " + labelMO.Tag.ToString() + " с ID:" + (labelMO.Tag as ModelObject).Identifier.ToString(), "Удаление", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    ModelObject mo = (labelMO.Tag as ModelObject);
                    mo.Select();
                    mo.Delete();
                }
            }
        }

        private void linkLabelMOh_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (labelMO.Tag is ModelObject)
            {
                ArrayList sel = new ArrayList();
                sel.Add(labelMO.Tag);
                Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
                MS.Select(sel);
                Tekla.Structures.ModelInternal.Operation.dotStartAction("ZoomToSelected", "");
                Tekla.Structures.ModelInternal.Operation.dotStartAction("HideUnselectedObjects", "");
            }
        }
 
        private void RunMacros(string comm)
        {
            string Name = GetMacroFileName();
            string MacrosPath = string.Empty;
            TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);
            if (MacrosPath.IndexOf(';')>0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';'));}
            File.WriteAllText(Path.Combine(MacrosPath, Name), "namespace Tekla.Technology.Akit.UserScript {public class Script {public static void Run(Tekla.Technology.Akit.IScript akit) {akit.ValueChange(\"main_frame\", \"" + comm + "\", \"1\");}}}");
			Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);			
        }
        private string GetMacroFileName()
        {
            lock (Random)
            {
                if (_TempFileIndex < 0)
                {
                    _TempFileIndex = Random.Next(0, MaxTempFiles);
                }
                else
                {
                    _TempFileIndex = (_TempFileIndex + 1) % MaxTempFiles;
                }

                return string.Format(FileNameFormat, _TempFileIndex);
            }
        }
        private static int _TempFileIndex = -1;
        private static readonly Random Random = new Random();
        private const int MaxTempFiles = 32;
        private const string FileNameFormat = "macro_{0:00}.cs";

        private void AssButton_Click_1(object sender, EventArgs e)
        {
            AssButton.Tag = !(bool)AssButton.Tag;
            if ((bool)AssButton.Tag)
            {
                AssButton.Text = "Search assemblies";                
                Text = "Search for assemlies";
                RunMacros("select_assemblies");
            }
            else
            {
                AssButton.Text = "Search parts ";
                Text = "Search for parts";
                RunMacros("sel_objects_in_joints");
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(linkLabel2.Text);
        }

        private void FindText_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar==(Char)13)
            {
                FindButF_Click(sender, null);
            }
        }

        private void OpenChert_Click(object sender, EventArgs e)
        {
            ModelObject Sel = null;
            TreeNode tn = treeView1.SelectedNode;
            if (tn.Tag is Assembly | tn.Tag is Part)
            {
                Sel = (ModelObject)tn.Tag;
            }
            else
            {
                int id = 0;
                int.TryParse(tn.Text.Substring(2),out id);
                if (id == 0) return;
                Sel = M.SelectModelObject(new Identifier(id));
            }            
            int DRAWING_ID = -1;
            Sel.GetReportProperty("DRAWING.ID", ref DRAWING_ID);
            if (DRAWING_ID != 0)
            {
                dr.DrawingHandler DH = new dr.DrawingHandler();
                dr.DrawingEnumerator DE = DH.GetDrawings();
                DE.MoveNext();
                dr.Drawing tDr = DE.Current;
                Type drType = tDr.GetType();
                sr.PropertyInfo propertyInfo = drType.GetProperty("Identifier", sr.BindingFlags.Instance | sr.BindingFlags.NonPublic);
                Identifier drIdentifier = new Identifier(DRAWING_ID);
                propertyInfo.SetValue(tDr, drIdentifier, null);
                Tekla.Structures.Model.Operations.Operation.DisplayPrompt("Открываем чертеж");
                DH.SetActiveDrawing(tDr, true);
            }
            else
            {
                Tekla.Structures.Model.Operations.Operation.DisplayPrompt("Чертеж не найден!!!");
            }

        }    
    }

    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            Application.EnableVisualStyles();
            Application.Run(new FindAssForm());
            string tv = Tekla.Structures.TeklaStructuresInfo.GetCurrentProgramVersion().Substring(0, 4);
            RegistryKey registryKey = Registry.CurrentUser.CreateSubKey("Software\\Tekla\\Structures\\" + tv + "\\Toolbars\\*UserToolbarZlatIng*");
            int butVal = registryKey.SubKeyCount;
            if ((string)registryKey.GetValue("TITLE") != "ZlatIng")
            {
                while (true)
                {
                    registryKey.SetValue("TITLE", "ZlatIng", RegistryValueKind.String);
                    registryKey.SetValue("ATTACH", 42, RegistryValueKind.DWord);
                    registryKey.SetValue("COLUMN", 1, RegistryValueKind.DWord);
                    registryKey.SetValue("DOCKED", 1, RegistryValueKind.DWord);
                    registryKey.SetValue("NR", 1, RegistryValueKind.DWord);
                    registryKey.SetValue("ROW", 1, RegistryValueKind.DWord);
                    registryKey.SetValue("FLOATINGWIDTH", 37, RegistryValueKind.DWord);
                    registryKey.SetValue("VISIBLE", (registryKey.Name.ToString().IndexOf("Drawing") > 1)?1:2, RegistryValueKind.DWord);
                    registryKey.SetValue("X", 100, RegistryValueKind.DWord);
                    registryKey.SetValue("Y", 100, RegistryValueKind.DWord);
                    if (registryKey.Name.ToString().IndexOf("Drawing") > 1)
                    {
                        break;
                    }
                    registryKey = Registry.CurrentUser.CreateSubKey("Software\\Tekla\\Structures\\" + tv + "\\Drawing\\Toolbars\\*UserToolbarZlatIng*");
                }
            }
                        
            int butSch = 1;
            bool MakroExist = false;
			RegistryKey butKey = null;
            while (butSch <= butVal)
            {
                butKey = Registry.CurrentUser.CreateSubKey("Software\\Tekla\\Structures\\" + tv + "\\Toolbars\\*UserToolbarZlatIng*\\" + butSch.ToString());
                if ((string)butKey.GetValue("ACTION") == "MacroGlobal_Поиск марок позиций")
                {
                    MakroExist = true;
                    break;
                }
                ++butSch;
            }
            if (!MakroExist)
            {                
                butKey = Registry.CurrentUser.CreateSubKey("Software\\Tekla\\Structures\\" + tv + "\\Toolbars\\*UserToolbarZlatIng*\\" + butSch.ToString());
                butKey.SetValue("ACTION", "MacroGlobal_Поиск марок позиций", RegistryValueKind.String);
                butKey.SetValue("TIP", "Поиск марок позиций", RegistryValueKind.String);
                butKey.SetValue("TYPE", "toolbutton", RegistryValueKind.String);
            }

        }
    }

}