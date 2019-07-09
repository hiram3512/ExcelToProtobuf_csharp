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

namespace HiProtobuf.UI
{
    public partial class HiProtobuf : Form
    {
        public HiProtobuf()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

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
        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = dialog.SelectedPath;
                Settings.Protoc_Path = textBox3.Text;
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            Settings.Protoc_Path = textBox3.Text;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = dialog.SelectedPath;
                Settings.Protobuf_Dll_Path = textBox4.Text;
            }
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            Settings.Protobuf_Dll_Path = textBox4.Text;
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
            Manager.Export();
        }
    }
}
