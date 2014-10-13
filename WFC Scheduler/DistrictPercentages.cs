using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WFC_Scheduler
{
    public partial class DistrictPercentages : Form
    {
        private List<District> currentDistrictList;

        public List<District> CurrentDistrictList
        {
            get { return currentDistrictList; }
            set { currentDistrictList = value; }
        }

        public int CanyonsCount
        {
            get { return Int32.Parse(canyonsTextBox.Text); } 
        }

        public int GraniteCount
        {
            get { return Int32.Parse(graniteTextBox.Text); }
        }

        public int JordanCount
        {
            get { return Int32.Parse(jordanTextBox.Text); }
        }

        public int MurrayCount
        {
            get { return Int32.Parse(murrayTextBox.Text); }
        }

        public int SaltLakeCount
        {
            get { return Int32.Parse(saltLakeTextBox.Text); }
        }

        public int TooeleCount
        {
            get { return Int32.Parse(tooeleTextBox.Text); }
        }

        public DistrictPercentages()
        {
            InitializeComponent();

            currentDistrictList = null;
            //populateFields();
        }

        private void checkTotals()
        {
            int sum = 0;
            foreach (Control currentBox in this.Controls)
            {
                foreach (District currentDist in currentDistrictList)
                {
                    if (currentBox is TextBox)
                    {
                        if (currentBox is TextBox && currentBox.Name.Equals(currentDist.DistrictName + "PercentBox"))
                        {
                            try
                            {
                                errorLabel.Visible = false;
                                Int32.Parse(currentBox.Text);
                                sum += Int32.Parse(currentBox.Text);
                            }
                            catch (FormatException)
                            {
                                errorLabel.Text = "Not a whole number";
                                errorLabel.Visible = true;
                            }
                        }


                    }
                }
            }

            if (sum != 100)
            {
                errorLabel.Text = "Total does not sum to 100";
                errorLabel.Visible = true;
            }
            else
            {
                if(currentDistrictList != null)
                {
                    foreach (District currentDist in currentDistrictList)
                    {
                        foreach (Control currentBox in this.Controls)
                        {
                            if (currentBox is TextBox && currentBox.Name.Equals(currentDist.DistrictName + "PercentBox"))
                            {
                                currentDist.DistrictPercentage = Int32.Parse(currentBox.Text);

                            }
                            else if (currentBox is TextBox && currentBox.Name.Equals(currentDist.DistrictName + "MaxBox"))
                            {
                                currentDist.MaxStudents = Int32.Parse(currentBox.Text);
                            }
                        }
                        currentDist.TotalStudents = 0;
                    }
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void changeDistrictTotalsButton_Click(object sender, EventArgs e)
        {
            checkTotals();
        }

        public void populateFields()
        {
            if (currentDistrictList != null)
            {
                foreach (Control currentControl in this.Controls)
                {
                    if (currentControl is Label)
                    {
                        currentControl.Visible = false;
                        //currentControl.Parent = null;
                    }
                    if (currentControl is TextBox)
                    {
                        currentControl.Visible = false;
                        //currentControl.Parent = null;
                    }
                }


                int x = 65;
                int y = 60;
                Point tempLoc = new Point(x,y);

                Label headingLabel = new Label();
                headingLabel.Font = new Font(headingLabel.Font.FontFamily.Name, 10, FontStyle.Bold);
                
                headingLabel.Location = tempLoc;
                headingLabel.Text = "District";
                headingLabel.Enabled = true;
                headingLabel.Visible = true;
                headingLabel.Parent = this;
                tempLoc.X += 100;

                Label percentLabel = new Label();
                percentLabel.Font = new Font(percentLabel.Font.FontFamily.Name, 10, FontStyle.Bold);

                percentLabel.Location = tempLoc;
                percentLabel.Text = "Percent";
                percentLabel.Enabled = true;
                percentLabel.Visible = true;
                percentLabel.Parent = this;
                tempLoc.X += 100;

                Label maxLabel = new Label();
                maxLabel.Font = new Font(headingLabel.Font.FontFamily.Name, 10, FontStyle.Bold);

                maxLabel.Location = tempLoc;
                maxLabel.Text = "Max Students";
                maxLabel.Enabled = true;
                maxLabel.Visible = true;
                maxLabel.Parent = this;
                
                tempLoc.X -= 200;
                tempLoc.Y += 30;


                #region place labels and textboxes on form for each district
                foreach (District temp in currentDistrictList)
                {
                    Label currLabel = new Label();
                    currLabel.Font = new Font(currLabel.Font.FontFamily.Name, 10);
                    currLabel.Name = temp.DistrictName + "Label";
                    currLabel.Location = tempLoc;
                    currLabel.Text = temp.DistrictName;
                    currLabel.Enabled = true;
                    currLabel.Visible = true;
                    currLabel.Parent = this;
                    tempLoc.X += 100;

                    TextBox currentPercentBox = new TextBox();
                    currentPercentBox.Name = temp.DistrictName + "PercentBox";
                    currentPercentBox.Width = 50;
                    currentPercentBox.Location = tempLoc;
                    currentPercentBox.Text = temp.DistrictPercentage.ToString();
                    currentPercentBox.Enabled = true;
                    currentPercentBox.Visible = true;
                    currentPercentBox.Parent = this;

                    tempLoc.X += 100;
                    TextBox currentMaxBox = new TextBox();
                    currentMaxBox.Name = temp.DistrictName + "MaxBox";
                    currentPercentBox.Width = 50;
                    currentMaxBox.Location = tempLoc;
                    currentMaxBox.Text = temp.MaxStudents.ToString();
                    currentMaxBox.Enabled = true;
                    currentMaxBox.Visible = true;
                    currentMaxBox.Parent = this;


                    tempLoc.X -= 200;
                    tempLoc.Y += 20;
                }
                #endregion

            }
            
        }
    }
}
