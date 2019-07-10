using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HiProtobuf.Lib;
using HiFramework.Log;

namespace HiProtobuf.UI
{
    public partial class HiProtobuf : Form
    {
        public HiProtobuf()
        {
            InitializeComponent();
            textBox1.Text = Settings.Export_Folder;
            textBox2.Text = Settings.Excel_Folder;
            textBox5.Text = Settings.Compiler_Path;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Log.OnPrint += (x) =>
            {
                textBox6.Text = Logger.Log;
            };
            Log.OnWarnning += (x) =>
            {
                textBox6.Text = Logger.Log;
            };
            Log.OnError += (x) =>
            {
                textBox6.Text = Logger.Log;
            };
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.SelectedPath;
                Settings.Export_Folder = textBox1.Text;
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Settings.Export_Folder = textBox1.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = dialog.SelectedPath;
                Settings.Excel_Folder = textBox2.Text;
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            Settings.Excel_Folder = textBox2.Text;
        }
        

        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = dialog.SelectedPath;
                Settings.Compiler_Path = textBox5.Text;
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            Settings.Compiler_Path = textBox5.Text;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Log.Print("开始导出");
            Manager.Export();
            Log.Print("导出结束");
            Config.Save();
        }
    }
}
