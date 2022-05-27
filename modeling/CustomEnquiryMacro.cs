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
	#region Form
    public class Form1 : Form
    {
		//ACTUAL WORK IS GOING HERE
		void Events_SelectionChangeEvent()
        {
            lock (_selectionEventHandlerLock)
            {
				Tekla.Structures.Model.UI.ModelObjectSelector selected = new Tekla.Structures.Model.UI.ModelObjectSelector();
				Tekla.Structures.Model.ModelObjectEnumerator manyO = (selected.GetSelectedObjects() as ModelObjectEnumerator);
				string txt = string.Empty;
				while (manyO.MoveNext())
				{
							string ID = string.Empty;
							string APOS = string.Empty;
							string PPOS = string.Empty;
							double LEN = 0;
							double WIDTH = 0;
							string UDA = string.Empty;
					if ((manyO.Current as Part) != null) //change for Assembly, Bolt, or whatever.
					{
						var pp = (manyO.Current as Part);

							pp.GetReportProperty("ASSEMBLY_POS", ref APOS);
							pp.GetReportProperty("PART_POS", ref PPOS);
							pp.GetReportProperty("LENGTH", ref LEN);
							pp.GetReportProperty("HEIGHT", ref WIDTH);
							pp.GetReportProperty("PART_POS", ref PPOS);
							pp.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.USER_PHASE", ref UDA);
							txt += APOS + "\t" + PPOS +"\t"+ Math.Round(LEN, 0) +"\t"+ Math.Round(WIDTH, 0) +"\t"+ UDA + Environment.NewLine;
					}
					else if ((manyO.Current as Assembly)!= null)
						{
							var pp = (manyO.Current as Assembly);

							pp.GetReportProperty("ID", ref ID);
							pp.GetReportProperty("ASSEMBLY_POS", ref APOS);
							pp.GetReportProperty("MAINPART.PART_POS", ref PPOS);
							pp.GetReportProperty("MAINPART.USERDEFINED.USER_PHASE", ref UDA);
							txt += ID + "/t" + APOS + "/t" + PPOS +"/t"+ UDA +Environment.NewLine;
						}
				}
				SetText(txt.ToString());
            }
        }

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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            //
            // textBox1
            //
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(0, 0);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(619, 187);
            this.textBox1.TabIndex = 1;
            //
            // Form1
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 189);
            this.Controls.Add(this.textBox1);
            this.Name = "Form1";
            this.Text = "Custom Inquiry Tool ";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
        #endregion
		#region FORM SETUP
        private System.Windows.Forms.TextBox textBox1;

        //create event first
        private Tekla.Structures.Model.Events _events = new Tekla.Structures.Model.Events();
        private object _selectionEventHandlerLock = new object();
        private Model TModel = new Model();

        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
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

        delegate void SetTextCallback(string text);

        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.textBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.textBox1.Text = text;
            }
        }
		#endregion
    }
    #endregion
	#region SCRIPT RUN
	    public class Script
        {
            public static void Run(Tekla.Technology.Akit.IScript akit)
            {
                Application.EnableVisualStyles();
                Application.Run(new Form1());
            }
        }
	#endregion
}