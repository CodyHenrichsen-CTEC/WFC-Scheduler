using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;

namespace WFC_Scheduler
{
    public partial class WorkshopGUIScreen : Form
    {

        private WorkshopData currentData;
        private string workshopDataPath, studentDataPath;
        private List<District> districtChangeList;

        public String StatusText
        {
            get { return statusLabel.Text;  }
            set { statusLabel.Text = value; }
        }

        public string WorkshopDataFilePath
        {
            get { return workshopDataPath; }
        }

        public string StudentDataFilePath
        {
            get { return studentDataPath; }
        }

        public DateTime selectedDate
        {
            get { return cutoffDateTimePicker.Value; }
        }


        public WorkshopGUIScreen()
        {
            InitializeComponent();
            currentData = new WorkshopData();
            districtChangeList = new List<District>();
            
        }

       
        private void workshopFileButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            
            
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;

            DialogResult userChoice = openFileDialog1.ShowDialog();
            if (userChoice == DialogResult.OK)
            {
                
                workshopDataPath = openFileDialog1.FileName;
                currentData.WorkshopInfoFilePath = WorkshopDataFilePath;
                cutoffDateTimePicker.Enabled = true;
                cutoffDateTimePicker.Value = DateTime.Today;
                
            }
            
        }

        private void districtPercentagesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DistrictPercentages currentDistrictTotals = new DistrictPercentages();
            if (currentData.DistrictList != null)
            {
                currentDistrictTotals.CurrentDistrictList = currentData.DistrictList;
                currentDistrictTotals.populateFields();

            }
            DialogResult setPercent = DialogResult.Abort;
            if (launchScheduleButton.Enabled && generateScheduleButton.Enabled)
            {
                setPercent = currentDistrictTotals.ShowDialog();
            }
            
            if (setPercent == DialogResult.OK)
            {

                if (currentDistrictTotals.CurrentDistrictList != null)
                {
                    currentData.DistrictList = currentDistrictTotals.CurrentDistrictList;
                    statusLabel.Text = "Percentages and/or Max changed \n click Generate Schedule to re-schedule";
                }
                else
                {
                    currentData.CanyonsPercent = currentDistrictTotals.CanyonsCount;
                    currentData.GranitePercent = currentDistrictTotals.GraniteCount;
                    currentData.JordanPercent = currentDistrictTotals.JordanCount;
                    currentData.MurrayPercent = currentDistrictTotals.MurrayCount;
                    currentData.SaltLakePercent = currentDistrictTotals.SaltLakeCount;
                    currentData.TooelePercent = currentDistrictTotals.TooeleCount;
                }

            }
        }

        private void chooseStudentRequestButton_Click(object sender, EventArgs e)
        {
            
            openFileDialog1.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);


            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;

            DialogResult userChoice = openFileDialog1.ShowDialog();
            if (userChoice == DialogResult.OK)
            {
                generateScheduleButton.Enabled = true;
                studentDataPath = openFileDialog1.FileName;
                currentData.RequestFilePath = StudentDataFilePath;
            }
        }

        private void generateScheduleButton_Click(object sender, EventArgs e)
        {
            currentData.startSchedule();
            statusLabel.Text = currentData.StatusText;
            currentData.verifyPopularity();
            statusLabel.Text = currentData.StatusText;
            currentData.generateScheduleFromLists();
            statusLabel.Text += currentData.StatusText;
            currentData.createWorkshopExportExcel();
            if (currentData.ScheduleOK)
            {
                launchScheduleButton.Enabled = true;
                launchScheduleButton.Visible = true;
            }
        }

        private void openSchedule()
        {
            string filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Workshop Final Schedule.xlsx";
            System.Diagnostics.Process.Start(filePath);
        }

        private void launchScheduleButton_Click(object sender, EventArgs e)
        {
            openSchedule();
        }

        private void cutoffDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            currentData.CutoffTime = cutoffDateTimePicker.Value;
            chooseStudentRequestButton.Enabled = true;
        }

        private void aboutSchedulerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About wfcAbout = new About();
            wfcAbout.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void helpFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFilesX86) + "\\Wasatch Front Consortium\\Scheduler\\workshop help file.pdf";
            System.Diagnostics.Process.Start(filePath);
        }

        
    }
}
