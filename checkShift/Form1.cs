using checkShift.Factory;
using checkShift.Models;
using System;
using System.Collections.Generic;
using System.Linq;
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
                if(!shiftFactory.check11Shift(personalShift, dateTimePicker1.Value, dateTimePicker2.Value, checkBox1.Checked, out errMsg))
                {
                    richTextBox1.AppendText(errMsg + "\r\n");
                }
            }

            foreach (PersonalShift personalShift in personalShifts)
            {
                if (!shiftFactory.check7Shift(personalShift, dateTimePicker1.Value, dateTimePicker2.Value, checkBox1.Checked, out errMsg))
                {
                    richTextBox1.AppendText(errMsg + "\r\n");
                }
            }

            foreach (PersonalShift personalShift in personalShifts)
            {
                if (!shiftFactory.check8Shift(personalShift, dateTimePicker1.Value, dateTimePicker2.Value, checkBox1.Checked, out errMsg))
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
            List<PersonalShift> personalShifts =  shiftFactory.ReadShirtFromDB(dateTimePicker1.Value, dateTimePicker2.Value, cmbUnit.Text);

            string errMsg = "";
            foreach (PersonalShift personalShift in personalShifts)
            {
                if (!shiftFactory.check11Shift(personalShift, dateTimePicker1.Value, dateTimePicker2.Value, checkBox1.Checked, out errMsg))
                {
                    richTextBox1.AppendText(errMsg + "\r\n");
                }
            }

            foreach (PersonalShift personalShift in personalShifts)
            {
                if (!shiftFactory.check7Shift(personalShift, dateTimePicker1.Value, dateTimePicker2.Value, checkBox1.Checked, out errMsg))
                {
                    richTextBox1.AppendText(errMsg + "\r\n");
                }
            }

            foreach (PersonalShift personalShift in personalShifts)
            {
                if (!shiftFactory.check8Shift(personalShift, dateTimePicker1.Value, dateTimePicker2.Value, checkBox1.Checked, out errMsg))
                {
                    
                    richTextBox1.AppendText(errMsg + "\r\n");
                }
            }

            MessageBox.Show("檢查完畢!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
            ShiftFactory shiftFactory = new ShiftFactory();

            List<ATTENDANCEDateTime> aTTENDANCEDateTimes = shiftFactory.getAttendace(dateTimePicker1.Value, dateTimePicker2.Value);

            List<ATTENDANCEDateTime> aTTENDANCEDates = aTTENDANCEDateTimes.Where(a=>a.OFFDATETIME < a.WORKDATETIME).ToList();

            foreach (ATTENDANCEDateTime aTTENDANCE in aTTENDANCEDates)
            {
                string errMsg = aTTENDANCE.TMNAME.Trim() + "(" + aTTENDANCE.KEYNO.Trim() + ") 出勤時間異常，上班時間:" + aTTENDANCE.WORKDATETIME + ",下班時間:" + aTTENDANCE.OFFDATETIME + "，請檢查!";
                richTextBox1.AppendText(errMsg + "\r\n");
            }

            aTTENDANCEDates  = aTTENDANCEDateTimes.Where(a => (a.OFFDATETIME - a.WORKDATETIME).Hours > 12).ToList();

            foreach (ATTENDANCEDateTime aTTENDANCE in aTTENDANCEDates)
            {
                string errMsg = aTTENDANCE.TMNAME.Trim() + "(" + aTTENDANCE.KEYNO.Trim() + ") 連續上班超過12小時，上班時間:" + aTTENDANCE.WORKDATETIME + ",下班時間:" + aTTENDANCE.OFFDATETIME + "，請檢查!";
                richTextBox1.AppendText(errMsg + "\r\n");
            }
            
            MessageBox.Show("檢查完畢!");
        }
    }
}
