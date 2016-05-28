using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Configuration;

namespace Scheduler_CPR
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        //Variables
        public string connString = "";

        //public int tempCount = 0;
        public double[] totalEmp1 = new double[7];
        public double[] totalEmp2 = new double[7];
        public double[] totalEmp3 = new double[7];
        public double[] totalEmp4 = new double[7];
        public double[] totalEmp5 = new double[7];
        public double[] totalEmp6 = new double[7];
        public double[] totalEmp7 = new double[7];
        public double[] totalEmp8 = new double[7];
        double[] totalMonday = new double[8];
        double[] totalTuesday = new double[8];
        double[] totalWednesday = new double[8];
        double[] totalThursday = new double[8];
        double[] totalFriday = new double[8];
        double[] totalSaturday = new double[8];
        double[] totalSunday = new double[8];

        //Population of Form
        private void populateTimeDropdown(ComboBox dropdown, DateTime startTime, DateTime endTime, TimeSpan interval)
        {
            dropdown.Items.Clear();

            DateTime time = startTime;

            while (time < endTime)
            {
                dropdown.Items.Add(time.ToString("hh:mm tt"));
                time = time.Add(interval);
            }
            dropdown.IntegralHeight = false;
        }
        public void populateStore(ComboBox cBox)
        {
            cBox.Items.Clear();

            DataSet myDataSet = new DataSet();
            OleDbConnection conn1 = null;
            string strAccessString = "SELECT * FROM Stores";
            try
            {
                conn1 = new OleDbConnection(connString);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection/n{0}", ex.Message);
                return;
            }

            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessString, conn1);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                conn1.Open();
                myDataAdapter.Fill(myDataSet, "Stores");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the database.\n{0}", ex.Message);
                return;
            }
            finally
            {
                conn1.Close();
            }

            DataRowCollection dra = myDataSet.Tables["Stores"].Rows;
            foreach (DataRow dr in dra)
            {
                Store st = new Store();

                st.storeID = Int32.Parse(dr[0].ToString());
                st.storeName = dr[1].ToString();

                cBox.Items.Add(st);
            }

            cBox.DisplayMember = "storeName";
            cBox.ValueMember = "storeID";
            cBox.IntegralHeight = false;
            if (conn1.State != ConnectionState.Closed) conn1.Close();
        }
        public void populateEmp(ComboBox cbox)
        {
            cbox.Items.Clear();

            DataSet myDataSet = new DataSet();
            OleDbConnection conn1 = null;
            string strAccessString = "SELECT * FROM Employees";

            try
            {

                conn1 = new OleDbConnection(connString);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to create a database connection/n{0}", ex.Message);
                return;
            }

            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessString, conn1);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                conn1.Open();
                myDataAdapter.Fill(myDataSet, "Employees");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Failed to retrieve the required data from the database.\n{0}", ex.Message);
                return;
            }
            finally
            {
                conn1.Close();
            }

            DataRowCollection dra = myDataSet.Tables["Employees"].Rows;
            foreach (DataRow dr in dra)
            {
                Employee empTemp = new Employee();
                empTemp.empID = Int32.Parse(dr[0].ToString());
                empTemp.empName = dr[1].ToString();
                empTemp.empEmail = dr[2].ToString();
                empTemp.empNotes = dr[3].ToString();
                cbox.Items.Add(empTemp);
            }


            cbox.DisplayMember = "empName";
            cbox.ValueMember = "empID";
            cbox.IntegralHeight = false;
            if (conn1.State != ConnectionState.Closed) conn1.Close();
        }
        public void populateDate()
        {
            cmbSelectDate.Items.Clear();
            DateTime dtMonday = new DateTime(2016, 5, 2);
            //cmbSelectDate.Items.Add(dtMonday.ToShortDateString());
            for(int i = 0; i<= 100; i++)
            {
                cmbSelectDate.Items.Add(dtMonday.AddDays(7 * i));
            }
            cmbSelectDate.IntegralHeight = false;
            cmbSelectDate.SelectedIndexChanged += new System.EventHandler(this.cmbSelectedDate_SelectedIndexChanged);
        }
        public void populateGroupBox(GroupBox gbx)
        {
            foreach (Control c in gbx.Controls)
            {
                try
                {
                    ComboBox cmb = (ComboBox)c;
                    populateTimeDropdown(cmb, new DateTime(2000, 1, 1, 9, 0, 0), new DateTime(2000, 1, 1, 21, 15, 0), new TimeSpan(0, 15, 0));
                }
                catch { }
            }
        }
        public void populateAll(bool option)
        {   
            if (!option) populateDate();
            if (!option) populateStore(cmbStore);
            foreach (Control c in this.Controls)
            {
                try
                {
                    GroupBox gbx = (GroupBox)c;
                    if (gbx.Name.Contains("Staff") && !option)
                    {
                        foreach (Control c2 in gbx.Controls)
                        {
                            try
                            {
                                ComboBox cmb = (ComboBox)c2;
                                populateEmp(cmb);
                                cmb.SelectedIndexChanged += new EventHandler(this.cmbEmployee_SelectedIndexChanged);
                            }
                            catch { }
                        }
                    }
                    else if (gbx.Name.Contains("Employee"))
                    {
                        foreach (Control c3 in gbx.Controls)
                        {
                            try
                            {
                                ComboBox cmb = (ComboBox)c3;
                                populateTimeDropdown(cmb, new DateTime(2000, 1, 1, 9, 0, 0), new DateTime(2000, 1, 1, 21, 15, 0), new TimeSpan(0, 15, 0));
                                cmb.SelectedIndexChanged += new EventHandler(this.cmbDay_SelectedIndexChanged);
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }
        }
        //public DataTable dTable()
        //{
        //    DateTime[] dtGroup = new DateTime[7];
        //    DataTable tbl = new DataTable();
        //    tbl.Columns.Add("Monday", typeof(DateTime));
        //    tbl.Columns.Add("Monday", typeof(DateTime));
        //    tbl.Columns.Add("Tuesday", typeof(DateTime));
        //    tbl.Columns.Add("Wednesday", typeof(DateTime));
        //    tbl.Columns.Add("Thursday", typeof(DateTime));
        //    tbl.Columns.Add("Friday", typeof(DateTime));
        //    tbl.Columns.Add("Saturday", typeof(DateTime));
        //    tbl.Columns.Add("Sunday", typeof(DateTime));
        //    tbl.Columns.Add("Employee Name", typeof(string));
        //    foreach(Control c in gbxDatesWeek1.Controls)
        //    {
        //        Label lbl = c as Label;

        //    }
        //}

        //Code that actually does stuff
        public void getTotals(String weekDay)
        {
            
            DateTime dt1 = DateTime.Now;
            DateTime dt2 = DateTime.Now;

            switch (weekDay)
            {
                case "monday":
                    {
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartMonday1.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndMonday1.SelectedItem.ToString());
                            totalMonday[0] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartMonday2.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndMonday2.SelectedItem.ToString());
                            totalMonday[1] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartMonday3.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndMonday3.SelectedItem.ToString());
                            totalMonday[2] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        { 
                            dt1 = DateTime.Parse(cmbStartMonday4.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndMonday4.SelectedItem.ToString());
                            totalMonday[3] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartMonday5.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndMonday5.SelectedItem.ToString());
                            totalMonday[4] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartMonday6.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndMonday6.SelectedItem.ToString());
                            totalMonday[5] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartMonday7.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndMonday7.SelectedItem.ToString());
                            totalMonday[6] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartMonday8.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndMonday8.SelectedItem.ToString());
                            totalMonday[7] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        txtTotalMonday.Text = (totalMonday[0] + totalMonday[1] + totalMonday[2] + totalMonday[3]).ToString("0.00");
                        txtTotalMonday2.Text = (totalMonday[4] + totalMonday[5] + totalMonday[6] + totalMonday[7]).ToString("0.00");
                        break;
                    }
                case "tuesday":
                    {
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartTuesday1.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndTuesday1.SelectedItem.ToString());
                            totalTuesday[0] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartTuesday2.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndTuesday2.SelectedItem.ToString());
                            totalTuesday[1] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartTuesday3.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndTuesday3.SelectedItem.ToString());
                            totalTuesday[2] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartTuesday4.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndTuesday4.SelectedItem.ToString());
                            totalTuesday[3] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartTuesday5.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndTuesday5.SelectedItem.ToString());
                            totalTuesday[4] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartTuesday6.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndTuesday6.SelectedItem.ToString());
                            totalTuesday[5] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartTuesday7.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndTuesday7.SelectedItem.ToString());
                            totalTuesday[6] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartTuesday8.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndTuesday8.SelectedItem.ToString());
                            totalTuesday[7] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        txtTotalTuesday.Text = (totalTuesday[0] + totalTuesday[1] + totalTuesday[2] + totalTuesday[3]).ToString("0.00");
                        txtTotalTuesday2.Text = (totalTuesday[4] + totalTuesday[5] + totalTuesday[6] + totalTuesday[7]).ToString("0.00");
                        break;
                    }
                case "wednesday":
                    {
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartWednesday1.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndWednesday1.SelectedItem.ToString());
                            totalWednesday[0] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartWednesday2.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndWednesday2.SelectedItem.ToString());
                            totalWednesday[1] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartWednesday3.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndWednesday3.SelectedItem.ToString());
                            totalWednesday[2] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartWednesday4.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndWednesday4.SelectedItem.ToString());
                            totalWednesday[3] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartWednesday5.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndWednesday5.SelectedItem.ToString());
                            totalWednesday[4] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartWednesday6.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndWednesday6.SelectedItem.ToString());
                            totalWednesday[5] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartWednesday7.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndWednesday7.SelectedItem.ToString());
                            totalWednesday[6] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartWednesday8.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndWednesday8.SelectedItem.ToString());
                            totalWednesday[7] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        txtTotalWednesday.Text = (totalWednesday[0] + totalWednesday[1] + totalWednesday[2] + totalWednesday[3]).ToString("0.00");
                        txtTotalWednesday2.Text = (totalWednesday[4] + totalWednesday[5] + totalWednesday[6] + totalWednesday[7]).ToString("0.00");
                        break;
                    }
                case "thursday":
                    {
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartThursday1.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndThursday1.SelectedItem.ToString());
                            totalThursday[0] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartThursday2.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndThursday2.SelectedItem.ToString());
                            totalThursday[1] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartThursday3.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndThursday3.SelectedItem.ToString());
                            totalThursday[2] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartThursday4.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndThursday4.SelectedItem.ToString());
                            totalThursday[3] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartThursday5.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndThursday5.SelectedItem.ToString());
                            totalThursday[4] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartThursday6.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndThursday6.SelectedItem.ToString());
                            totalThursday[5] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartThursday7.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndThursday7.SelectedItem.ToString());
                            totalThursday[6] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartThursday8.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndThursday8.SelectedItem.ToString());
                            totalThursday[7] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        txtTotalThursday.Text = (totalThursday[0] + totalThursday[1] + totalThursday[2] + totalThursday[3]).ToString("0.00");
                        txtTotalThursday2.Text = (totalThursday[4] + totalThursday[5] + totalThursday[6] + totalThursday[7]).ToString("0.00");
                        break;
                    }
                case "friday":
                    {
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartFriday1.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndFriday1.SelectedItem.ToString());
                            totalFriday[0] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartFriday2.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndFriday2.SelectedItem.ToString());
                            totalFriday[1] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartFriday3.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndFriday3.SelectedItem.ToString());
                            totalFriday[2] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartFriday4.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndFriday4.SelectedItem.ToString());
                            totalFriday[3] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartFriday5.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndFriday5.SelectedItem.ToString());
                            totalFriday[4] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartFriday6.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndFriday6.SelectedItem.ToString());
                            totalFriday[5] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartFriday7.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndFriday7.SelectedItem.ToString());
                            totalFriday[6] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartFriday8.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndFriday8.SelectedItem.ToString());
                            totalFriday[7] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        txtTotalFriday.Text = (totalFriday[0] + totalFriday[1] + totalFriday[2] + totalFriday[3]).ToString("0.00");
                        txtTotalFriday2.Text = (totalFriday[4] + totalFriday[5] + totalFriday[6] + totalFriday[7]).ToString("0.00");
                        break;
                    }
                case "saturday":
                    {
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSaturday1.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSaturday1.SelectedItem.ToString());
                            totalSaturday[0] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSaturday2.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSaturday2.SelectedItem.ToString());
                            totalSaturday[1] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSaturday3.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSaturday3.SelectedItem.ToString());
                            totalSaturday[2] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSaturday4.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSaturday4.SelectedItem.ToString());
                            totalSaturday[3] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSaturday5.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSaturday5.SelectedItem.ToString());
                            totalSaturday[4] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSaturday6.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSaturday6.SelectedItem.ToString());
                            totalSaturday[5] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSaturday7.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSaturday7.SelectedItem.ToString());
                            totalSaturday[6] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSaturday8.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSaturday8.SelectedItem.ToString());
                            totalSaturday[7] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        txtTotalSaturday.Text = (totalSaturday[0] + totalSaturday[1] + totalSaturday[2] + totalSaturday[3]).ToString("0.00");
                        txtTotalSaturday2.Text = (totalSaturday[4] + totalSaturday[5] + totalSaturday[6] + totalSaturday[7]).ToString("0.00");
                        break;
                    }
                case "sunday":
                    {
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSunday1.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSunday1.SelectedItem.ToString());
                            totalSunday[0] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSunday2.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSunday2.SelectedItem.ToString());
                            totalSunday[1] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSunday3.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSunday3.SelectedItem.ToString());
                            totalSunday[2] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSunday4.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSunday4.SelectedItem.ToString());
                            totalSunday[3] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSunday5.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSunday5.SelectedItem.ToString());
                            totalSunday[4] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSunday6.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSunday6.SelectedItem.ToString());
                            totalSunday[5] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSunday7.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSunday7.SelectedItem.ToString());
                            totalSunday[6] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        try
                        {
                            dt1 = DateTime.Parse(cmbStartSunday8.SelectedItem.ToString());
                            dt2 = DateTime.Parse(cmbEndSunday8.SelectedItem.ToString());
                            totalSunday[7] = (dt2 - dt1).TotalHours;
                        }
                        catch { }
                        txtTotalSunday.Text = (totalSunday[0] + totalSunday[1] + totalSunday[2] + totalSunday[3]).ToString("0.00");
                        txtTotalSunday2.Text = (totalSunday[4] + totalSunday[5] + totalSunday[6] + totalSunday[7]).ToString("0.00");
                        break;
                    }
            }
            double[] tdbl = new double[8];
            for (int i = 0; i <= 6; i++)
            {
                tdbl[i] = totalMonday[i] + totalTuesday[i] + totalWednesday[i] + totalThursday[i] + totalFriday[i] + totalSaturday[i] + totalSunday[i];
            }
            txtTotalEmp1.Text = tdbl[0].ToString("0.00");
            txtTotalEmp2.Text = tdbl[1].ToString("0.00");
            txtTotalEmp3.Text = tdbl[2].ToString("0.00");
            txtTotalEmp4.Text = tdbl[3].ToString("0.00");
            txtTotalEmp5.Text = tdbl[4].ToString("0.00");
            txtTotalEmp6.Text = tdbl[5].ToString("0.00");
            txtTotalEmp7.Text = tdbl[6].ToString("0.00");
            txtTotalEmp8.Text = tdbl[7].ToString("0.00");
            double week1Total = 0;
            double week2Total = 0;
            week1Total = tdbl[0] + tdbl[1] + tdbl[2] + tdbl[3];
            week2Total = tdbl[4] + tdbl[5] + tdbl[6] + tdbl[7];
            txtTotalWeek1.Text = week1Total.ToString("0.00");
            txtTotalWeek2.Text = week2Total.ToString("0.00");
        }
        public void resetAll(bool option)
        {   
            populateAll(false);
            gbxEmployee1.Enabled = false;
            gbxEmployee2.Enabled = false;
            gbxEmployee3.Enabled = false;
            gbxEmployee4.Enabled = false;
            gbxEmployee5.Enabled = false;
            gbxEmployee6.Enabled = false;
            gbxEmployee7.Enabled = false;
            gbxEmployee8.Enabled = false;
            gbxStaff.Enabled = false;
            cmbStore.Enabled = false;
            if (!option)
            {
                emptyTxt(false);
            }
            else emptyTxt(true);
        }
        public void emptyTxt(bool option)
        {
            foreach (Control c in this.Controls)
            {
                if (c.Name.Contains("Dates") && !option) 
                {
                    foreach (Control c1 in c.Controls)
                    {
                        try
                        {
                            Label lbl = (Label)c1;
                            lbl.Text = "";
                        }
                        catch { }
                    }
                }
                else if (c.Name.Contains("Totals"))
                {
                    foreach (Control c2 in c.Controls)
                    {
                        try
                        {
                            TextBox txt = (TextBox)c2;
                            txt.Text = "0.00";
                        }
                        catch { }
                    }
                }
            } 
            gbxWeekTotals.Enabled = false;
            gbxDayTotalsWeek1.Enabled = false;
        }
        public void checkConnString()
        {
            if (Settings.Default.dbasePath != null)
            {
                if (File.Exists(Settings.Default.dbasePath) != true)
                {
                    MessageBox.Show("Unable to locate database, please select new location.");
                    Settings.Default.dbasePath = GetFilePath();
                    Settings.Default.Save();
                }
            }
            else { Settings.Default.dbasePath = GetFilePath(); Settings.Default.Save(); }
            

            connString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + Settings.Default.dbasePath + "; Persist Security Info = True; Jet OLEDB:Database Password = M0bil3t3ch";
        }
        public void saveAll(string dateStart,string dateEnd)
        {
            StreamWriter sw;
            dateStart = dateStart.Replace("/", "_");
            dateEnd = dateEnd.Replace("/", "_");

            try
            {
                sw = new StreamWriter(dateStart + " through " + dateEnd + ".txt");
            }
            catch(IOException e)
            {
                MessageBox.Show(e.Message + "\n Cannot create file.");
                return;
            }
            try
            {
                sw.Write(cmbEmployee1.Text + "\r\nMonday ");
                sw.Write(lblMonday.Text + ": ");
                sw.Write(cmbStartMonday1.Text + " to " + cmbEndMonday1.Text);
                sw.Write("\r\nTuesday ");
                sw.Write(lblMonday.Text + ": ");
                sw.Write(cmbStartMonday1.Text + " to " + cmbEndTuesday1.Text);
                sw.Write("\r\nWednesday ");
                sw.Write(lblMonday.Text + ": ");
                sw.Write(cmbStartMonday1.Text + " to " + cmbEndWednesday1.Text);
                sw.Write("\r\nThursday ");
                sw.Write(lblMonday.Text + ": ");
                sw.Write(cmbStartMonday1.Text + " to " + cmbEndThursday1.Text);
                sw.Write("\r\nFriday ");
                sw.Write(lblMonday.Text + ": ");
                sw.Write(cmbStartMonday1.Text + " to " + cmbEndFriday1.Text);
                sw.Write("\r\nSaturday ");
                sw.Write(lblMonday.Text + ": ");
                sw.Write(cmbStartMonday1.Text + " to " + cmbEndSaturday1.Text);
                sw.Write("\r\nSunday ");
                sw.Write(lblMonday.Text + ": ");
                sw.Write(cmbStartMonday1.Text + " to " + cmbEndSunday1.Text);
            }
            catch(IOException e)
            {
                MessageBox.Show(e.Message + "\n Cannot write to file");
            }
            finally { sw.Close(); }
        }
        static string GetFilePath()
        {
            string returnValue = null;
            OpenFileDialog ofdData = new OpenFileDialog();
            ofdData.Filter = "Access Database (*.accdb) |*.accdb";
            ofdData.FilterIndex = 1;
            ofdData.Multiselect = false;
            if (ofdData.ShowDialog() == DialogResult.OK)
            {
                returnValue = ofdData.FileName;
            }
            else
            {
                MessageBox.Show("You Must Select The Database", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                returnValue = GetFilePath();
            }
            return returnValue;
        }

        //Main
        private void frmMain_Load(object sender, EventArgs e)
        {
            checkConnString();
            resetAll(false);
        }

        //Events
        private void cmbSelectedDate_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cbx = sender as ComboBox;
            DateTime d1 = (DateTime)cbx.SelectedItem;
            lblMonday.Text = (d1.ToString("MM/dd/yyyy"));
            lblTuesday.Text = (d1.AddDays(1).ToString("MM/dd/yyyy"));
            lblWednesday.Text = (d1.AddDays(2).ToString("MM/dd/yyyy"));
            lblThursday.Text = (d1.AddDays(3).ToString("MM/dd/yyyy"));
            lblFriday.Text = (d1.AddDays(4).ToString("MM/dd/yyyy"));
            lblSaturday.Text = (d1.AddDays(5).ToString("MM/dd/yyyy"));
            lblSunday.Text = (d1.AddDays(6).ToString("MM/dd/yyyy"));
            lblMonday2.Text = (d1.AddDays(7).ToString("MM/dd/yyyy"));
            lblTuesday2.Text = (d1.AddDays(8).ToString("MM/dd/yyyy"));
            lblWednesday2.Text = (d1.AddDays(9).ToString("MM/dd/yyyy"));
            lblThursday2.Text = (d1.AddDays(10).ToString("MM/dd/yyyy"));
            lblFriday2.Text = (d1.AddDays(11).ToString("MM/dd/yyyy"));
            lblSaturday2.Text = (d1.AddDays(12).ToString("MM/dd/yyyy"));
            lblSunday2.Text = (d1.AddDays(13).ToString("MM/dd/yyyy"));
            cmbStore.Enabled = true;
        }
        private void cmbStore_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbStore.SelectedIndex != -1)
            {
                gbxStaff.Enabled = true;
                emptyTxt(true);
            }
            populateAll(true);
        }
        private void cmbDay_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cmb = sender as ComboBox;
            string sendDay = "";
            string strTemp = cmb.Name;
            char[] numBer = new char[] { '1', '2', '3', '4', '5', '6', '7', '8' };
            if (strTemp.Contains("cmbStart") == true)
            {
                sendDay = strTemp.Replace("cmbStart", "");
            }
            else if (strTemp.Contains("cmbEnd") == true)
            {
                sendDay = strTemp.Replace("cmbEnd", "");
            }
            sendDay = sendDay.TrimEnd(numBer);
            getTotals(sendDay.ToLower());
        }
        private void cmbEmployee_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cbx = (ComboBox)sender;
            if (cbx.SelectedIndex != -1)
            {
                string strTemp = cbx.Name;
                switch (strTemp.ToLower())
                {
                    case "cmbemployee1":
                        gbxEmployee1.Enabled = true;
                        return;
                    case "cmbemployee2":
                        gbxEmployee2.Enabled = true;
                        return;
                    case "cmbemployee3":
                        gbxEmployee3.Enabled = true;
                        return;
                    case "cmbemployee4":
                        gbxEmployee4.Enabled = true;
                        return;
                    case "cmbemployee5":
                        gbxEmployee5.Enabled = true;
                        return;
                    case "cmbemployee6":
                        gbxEmployee6.Enabled = true;
                        return;
                    case "cmbemployee7":
                        gbxEmployee7.Enabled = true;
                        return;
                    case "cmbemployee8":
                        gbxEmployee8.Enabled = true;
                        return;
                }
            }

        }
        private void tsExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void resetToolStripMenuItem_Click(object sender, EventArgs e)
        {
                resetAll(false);
        }
    }
}