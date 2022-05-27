using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using Tekla.Structures;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using System.IO;
using Tekla.Technology.Akit.UserScript;
using Tekla.Technology.Scripting;
using System.Diagnostics;
using TSM = Tekla.Structures.Model;

namespace Tekla.Technology.Akit.UserScript
{
    #region FORM
    public class EventPannel_v2 : Form
    {
        #region FORM DESCRIPTION
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // textBox3
            // 
            this.textBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox3.Location = new System.Drawing.Point(0, 12);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox3.Size = new System.Drawing.Size(715, 38);
            this.textBox3.TabIndex = 5;
            this.textBox3.Text = "GUID;STANDARD;ASSEMBLY.MAINPART.USERDEFINED.ES_VD;ASSEMBLY.MAINPART.USERDEFINED.F" +
    "ILTER_KMD_DRAW_2";
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(0, 56);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(715, 318);
            this.dataGridView1.TabIndex = 14;
            // 
            // EventPannel_v2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(715, 374);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.textBox3);
            this.Name = "EventPannel_v2";
            this.Text = "CADSUPPORT: Custom Inquire Tool";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.EventPannel_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        #endregion

        //create event first 
        private Tekla.Structures.Model.Events _events = new Tekla.Structures.Model.Events();
        private object _selectionEventHandlerLock = new object();
        private TextBox textBox3;
        private Model TModel = new Model();

        public EventPannel_v2()
        {
            InitializeComponent();
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        }


        private void LoadSettings()
        {
            foreach (Control ct in this.Controls)
            {
                if (ct is TextBox)
                {
                    (ct as TextBox).Text = (String)Application.UserAppDataRegistry.GetValue((ct as TextBox).Name, String.Empty);
                }
            }
        }
        private void SaveSettings()
        {
            foreach (Control ct in this.Controls)
            {
                if (ct is TextBox)
                {
                    Application.UserAppDataRegistry.SetValue((ct as TextBox).Name, (ct as TextBox).Text);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadSettings();
            //add function to event 
            _events.SelectionChange += Events_SelectionChangeEvent;
            //and register it
            _events.Register();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //unregister it on exit.
            _events.UnRegister();
        }

        delegate void SetDataTableCallback(DataTable dataTable);

        private void SetDataSource(DataTable dataTable)
        {
            if (this.dataGridView1.InvokeRequired)
            {
                SetDataTableCallback d = new SetDataTableCallback(SetDataSource);
                this.Invoke(d, new object[] { dataTable });
            }
            else
            {
                this.dataGridView1.DataSource = dataTable;
            }
        }

        void Events_SelectionChangeEvent()
        {
            //Make sure that the inner code block is running synchronously.
            lock (_selectionEventHandlerLock)
            {
                //Arrange DataTable
                string[] PropToGet = textBox3.Text.Split(';');
                DataTable datatable = new System.Data.DataTable("Specification");
                foreach (string st in PropToGet) datatable.Columns.Add(st);
                int i = 0;

                //Get Model Objects from tekla
                Tekla.Structures.Model.UI.ModelObjectSelector selected = new Tekla.Structures.Model.UI.ModelObjectSelector();
                Tekla.Structures.Model.ModelObjectEnumerator manyO = (selected.GetSelectedObjects() as ModelObjectEnumerator);
                string txt = string.Empty;

                while (manyO.MoveNext())
                {
                    datatable.Rows.Add();
                    if ((manyO.Current as ModelObject) != null)
                    {
                        var pp = manyO.Current as ModelObject;
                        foreach (string str in PropToGet)
                        {
                            string get_prop_string = "";
                            double get_prop_double = 0;
                            int get_prop_int = 0;
                            pp.GetReportProperty(str, ref get_prop_string);
                            if (get_prop_string == "")
                            {
                                pp.GetReportProperty(str, ref get_prop_double);
                                if (get_prop_double == 0)
                                {
                                    pp.GetReportProperty(str, ref get_prop_int);
                                    if (get_prop_int != 0) datatable.Rows[i][str] = get_prop_int.ToString();
                                }
                                else
                                {
                                    datatable.Rows[i][str] = get_prop_double.ToString();
                                }
                            }
                            else
                            {
                                datatable.Rows[i][str] = get_prop_string;
                            }
                        }
                    }
                    i++;
                }
                SetDataSource(datatable);
            }
        }
        private void EventPannel_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        //private List<MyTeklaModelObject> myList { get; set; }

        private DataGridView dataGridView1;
    }
    #endregion
	#region SCRIPT RUN
	    public class Script
        {
            public static void Run(Tekla.Technology.Akit.IScript akit)
            {
                Application.EnableVisualStyles();
                Application.Run(new EventPannel_v2());
            }
        }
	#endregion
}