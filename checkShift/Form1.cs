using checkShift.Factory;
using checkShift.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;

using System.Windows.Forms;


namespace checkShift
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
            ShiftFactory shiftFactory = new ShiftFactory();
            if(textBox1.Text == "")
            {
                MessageBox.Show("請先選擇班表!");
                return;
            }
            List<PersonalShift> personalShifts = shiftFactory.ReadShirt(textBox1.Text, dateTimePicker1.Value, dateTimePicker2.Value);

            string errMsg = "";
            foreach (PersonalShift personalShift in personalShifts)
            {
                if(!shiftFactory.checkShift(personalShift, dateTimePicker1.Value, dateTimePicker2.Value, checkBox1.Checked, out errMsg))
                {
                    richTextBox1.AppendText(errMsg + "\r\n");
                }
            }

            MessageBox.Show("檢查完畢!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx;*.xls";
            openFileDialog1.FileName = "選擇班表";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox1.Text = openFileDialog1.FileName;
                }
                catch 
                {
                  
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
            ShiftFactory shiftFactory = new ShiftFactory();
            List<PersonalShift> personalShifts = shiftFactory.ReadShirtFromDB(dateTimePicker1.Value, dateTimePicker2.Value);

            string errMsg = "";
            foreach (PersonalShift personalShift in personalShifts)
            {
                if (!shiftFactory.checkShift(personalShift, dateTimePicker1.Value, dateTimePicker2.Value, checkBox1.Checked, out errMsg))
                {
                    richTextBox1.AppendText(errMsg + "\r\n");
                }
            }

            MessageBox.Show("檢查完畢!");
        }
    }
}
