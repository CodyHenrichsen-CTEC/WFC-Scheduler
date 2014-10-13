using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Data;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Collections;
using System.IO;

namespace WFC_Scheduler
{
    class WorkshopTableData
    {

        private int canyonsPercent, granitePercent, murrayPercent, jordanPercent, saltLakePercent, tooelePercent, districtTotals;

        private string requestFilePath, workshopInfoFilePath, statusText, failedCutoff, exceedsCount;
        private DateTime cutoffTime;
        private DataTable roomTable, schoolTable, districtTable, presenterTable, requestTable, studentTable, sessionTable;
        private Dictionary<String, Int32> PopularClasses, FourthClasses, ThirdClasses, SecondClasses, FirstClasses;


        private bool checkSheetsMatch(ExcelWorkbook current)
        {
            bool sheetsMatch = false;
            int rooms = 0, schools = 0, districts = 0, presenters = 0;

            for (int count = 0; count < 4; count++)
            {
                ExcelWorksheet currentWorksheet = current.Worksheets[count + 1];
                if (currentWorksheet.Name.Equals("Rooms"))
                {
                    rooms = 1;
                }
                if (currentWorksheet.Name.Equals("Schools"))
                {
                    schools = 1;
                }
                if (currentWorksheet.Name.Equals("Districts"))
                {
                    districts = 1;
                }
                if (currentWorksheet.Name.Equals("Presenters"))
                {
                    presenters = 1;
                }
            }
            sheetsMatch = (rooms + schools + districts + presenters == 4);
            return sheetsMatch;

        }

        #region Table Extraction version

        /// <summary>
        /// Check the popularity using datatables
        /// </summary>
        private void checkPopularity()
        {
            PopularClasses = new Dictionary<String, int>();
            FourthClasses = new Dictionary<String, int>();
            ThirdClasses = new Dictionary<String, int>();
            SecondClasses = new Dictionary<String, int>();
            FirstClasses = new Dictionary<String, int>();

            foreach (DataRow presenterRow in presenterTable.Rows)
            {
                foreach (DataRow requestRow in requestTable.Rows)
                {
                    #region Loop over all the requests and populate dictionaries
                    String currentPresenter = (String)presenterRow["PresenterTitle"];

                    if (((String)requestRow["Student Request 1"]).Equals(currentPresenter) || ((String)requestRow["Student Request 2"]).Equals(currentPresenter) || ((String)requestRow["Student Request 3"]).Equals(currentPresenter) || ((String)requestRow["Student Request 4"]).Equals(currentPresenter) || ((String)requestRow["Student Request 5"]).Equals(currentPresenter))
                    {
                        if (PopularClasses.ContainsKey(currentPresenter))
                        {
                            int currentCount = PopularClasses[currentPresenter];
                            ++currentCount;
                            PopularClasses[currentPresenter] = currentCount;
                        }
                        else
                        {
                            PopularClasses.Add(currentPresenter, 1);
                        }
                    }
                    if (((String)requestRow["Student Request 1"]).Equals(currentPresenter) || ((String)requestRow["Student Request 2"]).Equals(currentPresenter) || ((String)requestRow["Student Request 3"]).Equals(currentPresenter) || ((String)requestRow["Student Request 4"]).Equals(currentPresenter))
                    {
                        if (FourthClasses.ContainsKey(currentPresenter))
                        {
                            FourthClasses[currentPresenter] = ++FourthClasses[currentPresenter];
                        }
                        else
                        {
                            FourthClasses.Add(currentPresenter, 1);
                        }
                    }
                    if (((String)requestRow["Student Request 1"]).Equals(currentPresenter) || ((String)requestRow["Student Request 2"]).Equals(currentPresenter) || ((String)requestRow["Student Request 3"]).Equals(currentPresenter))
                    {
                        if (ThirdClasses.ContainsKey(currentPresenter))
                        {
                            ThirdClasses[currentPresenter] = ++ThirdClasses[currentPresenter];
                        }
                        else
                        {
                            ThirdClasses.Add(currentPresenter, 1);
                        }
                    }
                    if (((String)requestRow["Student Request 1"]).Equals(currentPresenter) || ((String)requestRow["Student Request 2"]).Equals(currentPresenter))
                    {
                        if (SecondClasses.ContainsKey(currentPresenter))
                        {
                            SecondClasses[currentPresenter] = ++SecondClasses[currentPresenter];
                        }
                        else
                        {
                            SecondClasses.Add(currentPresenter, 1);
                        }
                    }
                    if (((String)requestRow["Student Request 1"]).Equals(currentPresenter))
                    {
                        if (FirstClasses.ContainsKey(currentPresenter))
                        {
                            FirstClasses[currentPresenter] = ++FirstClasses[currentPresenter];
                        }
                        else
                        {
                            FirstClasses.Add(currentPresenter, 1);
                        }
                    }
                    #endregion
                }
            }
            PopularClasses = PopularClasses.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
        }


        private void importExcelToTables()
        {

            FileInfo workshopFile = new FileInfo(workshopInfoFilePath);
            using (ExcelPackage currentExcelFile = new ExcelPackage(workshopFile))
            {
                ExcelWorkbook currentWorkbook = currentExcelFile.Workbook;
                if (currentWorkbook != null)
                {

                    if (!checkSheetsMatch(currentWorkbook))
                    {
                        statusText = "Sheets do not have correct names upload a new file";
                    }
                    else
                    {
                        ExcelWorksheet roomSheet = currentWorkbook.Worksheets["Rooms"];
                        ExcelWorksheet districtSheet = currentWorkbook.Worksheets["Districts"];
                        ExcelWorksheet presentersSheet = currentWorkbook.Worksheets["Presenters"];
                        ExcelWorksheet schoolsSheet = currentWorkbook.Worksheets["Schools"];
                        addToRoomTable(roomSheet);
                        addToPresenterTable(presentersSheet);
                        addToDistrictTable(districtSheet);
                        addToSchoolTable(schoolsSheet);
                    }


                }
            }

            FileInfo requestFile = new FileInfo(requestFilePath);
            using (ExcelPackage currentExcelFile = new ExcelPackage(requestFile))
            {
                ExcelWorkbook currentWorkbook = currentExcelFile.Workbook;
                if (currentWorkbook != null)
                {
                    ExcelWorksheet requestSheet = currentWorkbook.Worksheets[1];
                    addToRequestTable(requestSheet);
                }

            }

            createStudentTable();
            createSessionTable();

        }

        private void addToRequestTable(ExcelWorksheet requestSheet)
        {
            requestTable = new DataTable();

            string firstName, lastName, email, phone, requestOne, requestTwo, requestThree, requestFour, requestFive, school, studentGradeLevel, studentTeacher, studentNumber;
            int studentID;
            DateTime studentTime;


            requestTable.Columns.Add(new DataColumn("Student ID", typeof(int)));
            requestTable.Columns.Add(new DataColumn("Student Time", typeof(DateTime)));
            requestTable.Columns.Add(new DataColumn("Student First Name", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Last Name", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Email", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Phone", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student School", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Request 1", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Request 2", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Request 3", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Request 4", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Request 5", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Grade Level", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Number", typeof(String)));
            requestTable.Columns.Add(new DataColumn("Student Teacher Name", typeof(String)));

            requestTable.PrimaryKey = new DataColumn[] { requestTable.Columns["Student ID"] };

            for (int row = 2; row <= requestSheet.Dimension.End.Row; row++)
            {
                studentID = row - 1;

                double serialDate = double.Parse(requestSheet.Cells[row, 1].Value.ToString());
                studentTime = DateTime.FromOADate(serialDate);
                firstName = (String)requestSheet.Cells[row, 2].Value.ToString();
                lastName = (String)requestSheet.Cells[row, 3].Value.ToString();
                email = (String)requestSheet.Cells[row, 4].Value.ToString();
                phone = (String)requestSheet.Cells[row, 5].Value.ToString();
                school = (String)requestSheet.Cells[row, 6].Value.ToString();
                requestOne = (String)requestSheet.Cells[row, 7].Value.ToString();
                requestTwo = (String)requestSheet.Cells[row, 8].Value.ToString();
                requestThree = (String)requestSheet.Cells[row, 9].Value.ToString();
                requestFour = (String)requestSheet.Cells[row, 10].Value.ToString();
                requestFive = (String)requestSheet.Cells[row, 11].Value.ToString();
                studentGradeLevel = (String)requestSheet.Cells[row, 12].Value.ToString();
                studentNumber = (String)requestSheet.Cells[row, 13].Value.ToString();
                studentTeacher = (String)requestSheet.Cells[row, 14].Value.ToString();

                DataRow currentRequest = requestTable.NewRow();
                currentRequest[0] = studentID;
                currentRequest[1] = studentTime;
                currentRequest[2] = firstName;
                currentRequest[3] = lastName;
                currentRequest[4] = email;
                currentRequest[5] = phone;
                currentRequest[6] = school;
                currentRequest[7] = requestOne;
                currentRequest[8] = requestTwo;
                currentRequest[9] = requestThree;
                currentRequest[10] = requestFour;
                currentRequest[11] = requestFive;
                currentRequest[12] = studentGradeLevel;
                currentRequest[13] = studentNumber;
                currentRequest[14] = studentTeacher;

                requestTable.Rows.Add(currentRequest);
            }
        }

        private void addToSchoolTable(ExcelWorksheet schoolsSheet)
        {
            schoolTable = new DataTable();
            schoolTable.Columns.Add(new DataColumn("SchoolName", typeof(String)));
            schoolTable.Columns.Add(new DataColumn("School District", typeof(String)));
            schoolTable.PrimaryKey = new DataColumn[] { schoolTable.Columns["SchoolName"] };

            String schoolName, districtName;

            for (int row = 2; row <= schoolsSheet.Dimension.End.Row; row++)
            {
                schoolName = (String)schoolsSheet.Cells[row, 1].Value;
                districtName = (String)schoolsSheet.Cells[row, 2].Value;
                DataRow currentSchool = schoolTable.NewRow();
                currentSchool[0] = schoolName;
                currentSchool[1] = districtName;
                schoolTable.Rows.Add(currentSchool);
            }
        }

        private void addToDistrictTable(ExcelWorksheet districtSheet)
        {
            districtTable = new DataTable();

            districtTable.Columns.Add(new DataColumn("District ID", typeof(int)));
            districtTable.Columns.Add(new DataColumn("District Name", typeof(String)));
            districtTable.Columns.Add(new DataColumn("District Percentage", typeof(int)));

            districtTable.PrimaryKey = new DataColumn[] { districtTable.Columns["District Name"] };

            String districtName;
            int districtPercentage;

            for (int row = 2; row <= districtSheet.Dimension.End.Row; row++)
            {
                districtName = (String)districtSheet.Cells[row, 1].Value;
                districtPercentage = Convert.ToInt32(districtSheet.Cells[row, 2].Value);

                DataRow currentDistrict = districtTable.NewRow();
                currentDistrict[0] = (row - 1);
                currentDistrict[1] = districtName;
                currentDistrict[2] = districtPercentage;


                districtTable.Rows.Add(currentDistrict);
            }
        }

        private void addToPresenterTable(ExcelWorksheet presentersSheet)
        {
            presenterTable = new DataTable();
            presenterTable.Columns.Add(new DataColumn("PresenterTitle", typeof(String)));
            presenterTable.Columns.Add(new DataColumn("PresenterDescription", typeof(String)));
            presenterTable.Columns.Add(new DataColumn("PresenterRoom", typeof(String)));
            presenterTable.PrimaryKey = new DataColumn[] { presenterTable.Columns["PresenterTitle"] };
            //presenterTable[2] //search for me!
            String PresenterTitle, PresenterDescription, presenterRoom;
            //int roomID;

            for (int row = 2; row <= presentersSheet.Dimension.End.Row; row++)
            {
                PresenterTitle = (String)presentersSheet.Cells[row, 1].Value;
                PresenterDescription = (String)presentersSheet.Cells[row, 2].Value;
                presenterRoom = (String)presentersSheet.Cells[row, 3].Value;
                //roomID = (Int32)roomTable.Rows.Find((String)presentersSheet.Cells[row, 3].Value)["RoomID"];

                DataRow currentPresenter = presenterTable.NewRow();
                currentPresenter[0] = PresenterTitle;
                currentPresenter[1] = PresenterDescription;
                currentPresenter[2] = presenterRoom;

                presenterTable.Rows.Add(currentPresenter);
            }
        }

        private void addToRoomTable(ExcelWorksheet roomSheet)
        {
            roomTable = new DataTable();
            roomTable.Columns.Add(new DataColumn("RoomID", typeof(int)));
            roomTable.Columns.Add(new DataColumn("RoomName", typeof(String)));
            roomTable.Columns.Add(new DataColumn("RoomCapacity", typeof(int)));

            roomTable.PrimaryKey = new DataColumn[] { roomTable.Columns["RoomName"] };



            String roomName;
            int roomCapacity;

            for (int row = 2; row <= roomSheet.Dimension.End.Row; row++)
            {
                roomName = (String)roomSheet.Cells[row, 1].Value;
                roomCapacity = Convert.ToInt32(roomSheet.Cells[row, 2].Value);

                DataRow currentRoom = roomTable.NewRow();
                currentRoom[0] = (row - 1);
                currentRoom[1] = roomName;
                currentRoom[2] = roomCapacity;

                roomTable.Rows.Add(currentRoom);
            }
        }

        private void createStudentTable()
        {
            studentTable = new DataTable();

            studentTable.Columns.Add(new DataColumn("Student ID", typeof(int)));
            studentTable.Columns.Add(new DataColumn("Student First Name", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Last Name", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student School", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Email", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Phone", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Grade Level", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Number", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Teacher Name", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Session A", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Session B", typeof(String)));
            studentTable.Columns.Add(new DataColumn("Student Session C", typeof(String)));

            studentTable.PrimaryKey = new DataColumn[] { studentTable.Columns["Student ID"] };



            foreach (DataRow currentRequest in requestTable.Rows)
            {
                DataRow currentStudent = studentTable.NewRow();

                currentStudent[0] = currentRequest[0];  //id
                currentStudent[1] = currentRequest[2];  //first
                currentStudent[2] = currentRequest[3];  //last
                currentStudent[3] = currentRequest[4];  //school
                currentStudent[4] = currentRequest[5];  //email
                currentStudent[5] = currentRequest[6];  //phone
                currentStudent[6] = currentRequest[12]; //grade
                currentStudent[7] = currentRequest[13]; //number
                currentStudent[8] = currentRequest[14]; //teacher

                currentStudent[9] = "0";                //default 1
                currentStudent[10] = "0";               //default 2
                currentStudent[11] = "0";               //default 3

                studentTable.Rows.Add(currentStudent);
            }

        }

        private void createSessionTable()
        {
            sessionTable = new DataTable();

            sessionTable.Columns.Add(new DataColumn("Session ID", typeof(int)));
            sessionTable.Columns.Add(new DataColumn("Session Room", typeof(String)));
            sessionTable.Columns.Add(new DataColumn("SessionPresenter", typeof(String)));
            sessionTable.Columns.Add(new DataColumn("Session A Count", typeof(int)));
            sessionTable.Columns.Add(new DataColumn("Session B Count", typeof(int)));
            sessionTable.Columns.Add(new DataColumn("Session C Count", typeof(int)));
            sessionTable.PrimaryKey = new DataColumn[] { sessionTable.Columns["SessionPresenter"] };

            for (int rowCount = 1; rowCount < presenterTable.Rows.Count; rowCount++)
            {
                DataRow currentRow = sessionTable.NewRow();
                currentRow[0] = rowCount;
                currentRow[1] = presenterTable.Rows[rowCount][2];
                currentRow[2] = presenterTable.Rows[rowCount][0];
                currentRow[3] = 0;
                currentRow[4] = 0;
                currentRow[5] = 0;

                sessionTable.Rows.Add(currentRow);
            }
        }

        private void generateSchedule()
        {


            foreach (KeyValuePair<string, int> classIDandCount in PopularClasses)
            {
                String searchValue = "[SessionPresenter] = '" + classIDandCount.Key + "'";
                DataRow[] otherTest = sessionTable.Select("SessionPresenter LIKE '" + classIDandCount.Key + "'");
                DataRow testRow = sessionTable.Rows.Find(classIDandCount.Key);
                DataRow[] sessionRows = sessionTable.Select(searchValue);

                Dictionary<String, Int32> currentDistrictCount = new Dictionary<String, int>();

                foreach (DataRow currentDistrict in districtTable.Rows)
                {
                    currentDistrictCount.Add((String)currentDistrict["District Name"], 0);
                }

                if (sessionRows.Count() == 0)
                {
                    continue;
                }
                DataRow sessionRow = sessionRows[0];

                DataRow roomRow = roomTable.Rows.Find(sessionRow["Session Room"]);

                int totalRoomCapacity = 3 * (Int32)roomRow["RoomCapacity"];
                int currentCapacity = (Int32)roomRow["RoomCapacity"];

                int currentSessionACount = 0;
                int currentSessionBCount = 0;
                int currentSessionCCount = 0;

                foreach (DataRow currentRequest in requestTable.Rows)
                {
                    if ((DateTime)currentRequest["Student Time"] > cutoffTime)
                    {
                        //set each session for associated student name to "no sessions, enrolled after cutoff time"
                        String failedCutoff = "no sessions, enrolled after cutoff time";
                        studentTable.Rows.Find(currentRequest["Student ID"])["Student Session A"] = failedCutoff;
                        studentTable.Rows.Find(currentRequest["Student ID"])["Student Session B"] = failedCutoff;
                        studentTable.Rows.Find(currentRequest["Student ID"])["Student Session C"] = failedCutoff;

                    }

                    else
                    #region Student can be scheduled
                    {
                        int currentStudentID = (Int32)currentRequest["Student ID"];

                        string currentSchoolName = (String)requestTable.Rows.Find(currentStudentID)["Student School"];
                        string currentDistrictName = (String)schoolTable.Rows.Find(currentSchoolName)["School District"];
                        int currentPercentage = (Int32)districtTable.Rows.Find(currentDistrictName)["District Percentage"];

                        double currentPerc = currentPercentage / 100;
                        int currentDistrictMax = (Int32)(((Double)(currentPercentage / 100.00) * (Double)totalRoomCapacity));

                        #region Fits in all
                        if (classIDandCount.Value < totalRoomCapacity)
                        {
                            if (((String)currentRequest["Student Request 1"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 2"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 3"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 4"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 5"]).Equals(classIDandCount.Key))
                            {
                                DataRow currentStudentRow = studentTable.Rows.Find(currentStudentID);
                                if ((((String)currentStudentRow["Student Session A"]).Equals("0") && currentSessionACount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session A"] = classIDandCount.Key;
                                    currentSessionACount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session B"]).Equals("0") && currentSessionBCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session B"] = classIDandCount.Key;
                                    currentSessionBCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session C"]).Equals("0") && currentSessionCCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session C"] = classIDandCount.Key;
                                    currentSessionCCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                            }
                        }
                        #endregion
                        #region Fourth and better requests
                        else if (!FourthClasses.ContainsKey(classIDandCount.Key))
                        {
                            continue;
                        }
                        else if (FourthClasses[classIDandCount.Key] < totalRoomCapacity && FourthClasses.ContainsKey(classIDandCount.Key))
                        {
                            if (((String)currentRequest["Student Request 1"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 2"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 3"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 4"]).Equals(classIDandCount.Key))
                            {
                                DataRow currentStudentRow = studentTable.Rows.Find(currentStudentID);
                                if ((((String)currentStudentRow["Student Session A"]).Equals("0") && currentSessionACount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session A"] = classIDandCount.Key;
                                    currentSessionACount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session B"]).Equals("0") && currentSessionBCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session B"] = classIDandCount.Key;
                                    currentSessionBCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session C"]).Equals("0") && currentSessionCCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session C"] = classIDandCount.Key;
                                    currentSessionCCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                            }
                        }
                        #endregion
                        #region Third and better requests only
                        else if (!ThirdClasses.ContainsKey(classIDandCount.Key))
                        {
                            continue;
                        }

                        else if (ThirdClasses[classIDandCount.Key] < totalRoomCapacity && ThirdClasses.ContainsKey(classIDandCount.Key))
                        {
                            if (((String)currentRequest["Student Request 1"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 2"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 3"]).Equals(classIDandCount.Key))
                            {
                                DataRow currentStudentRow = studentTable.Rows.Find(currentStudentID);
                                if ((((String)currentStudentRow["Student Session A"]).Equals("0") && currentSessionACount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session A"] = classIDandCount.Key;
                                    currentSessionACount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session B"]).Equals("0") && currentSessionBCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session B"] = classIDandCount.Key;
                                    currentSessionBCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session C"]).Equals("0") && currentSessionCCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session C"] = classIDandCount.Key;
                                    currentSessionCCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                            }
                        }
                        #endregion
                        #region Second and better requests only
                        else if (!SecondClasses.ContainsKey(classIDandCount.Key))
                        {
                            continue;
                        }
                        else if (SecondClasses[classIDandCount.Key] < totalRoomCapacity && SecondClasses.ContainsKey(classIDandCount.Key))
                        {
                            if (((String)currentRequest["Student Request 1"]).Equals(classIDandCount.Key) || ((String)currentRequest["Student Request 2"]).Equals(classIDandCount.Key))
                            {
                                DataRow currentStudentRow = studentTable.Rows.Find(currentStudentID);
                                if ((((String)currentStudentRow["Student Session A"]).Equals("0") && currentSessionACount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session A"] = classIDandCount.Key;
                                    currentSessionACount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session B"]).Equals("0") && currentSessionBCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session B"] = classIDandCount.Key;
                                    currentSessionBCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session C"]).Equals("0") && currentSessionCCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session C"] = classIDandCount.Key;
                                    currentSessionCCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                            }
                        }
                        #endregion
                        #region First choices only
                        else if (!FirstClasses.ContainsKey(classIDandCount.Key))
                        {
                            continue;
                        }

                        else if (FirstClasses[classIDandCount.Key] < totalRoomCapacity && FirstClasses.ContainsKey(classIDandCount.Key))
                        {
                            if (((String)currentRequest["Student Request 1"]).Equals(classIDandCount.Key))
                            {
                                DataRow currentStudentRow = studentTable.Rows.Find(currentStudentID);
                                if ((((String)currentStudentRow["Student Session A"]).Equals("0") && currentSessionACount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session A"] = classIDandCount.Key;
                                    currentSessionACount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session B"]).Equals("0") && currentSessionBCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session B"] = classIDandCount.Key;
                                    currentSessionBCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session C"]).Equals("0") && currentSessionCCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session C"] = classIDandCount.Key;
                                    currentSessionCCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                            }
                        }
                        #endregion
                        #region Less than first choice totals
                        else
                        {
                            if (((String)currentRequest["Student Request 1"]).Equals(classIDandCount.Key))
                            {
                                DataRow currentStudentRow = studentTable.Rows.Find(currentStudentID);
                                if (currentStudentRow == null)
                                {
                                    continue;
                                }

                                if ((((String)currentStudentRow["Student Session A"]).Equals("0") && currentSessionACount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session A"] = classIDandCount.Key;
                                    currentSessionACount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session B"]).Equals("0") && currentSessionBCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session B"] = classIDandCount.Key;
                                    currentSessionBCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                                else if ((((String)currentStudentRow["Student Session C"]).Equals("0") && currentSessionCCount < currentCapacity) && currentDistrictCount[currentDistrictName] < currentDistrictMax)
                                {
                                    currentStudentRow["Student Session C"] = classIDandCount.Key;
                                    currentSessionCCount++;
                                    currentDistrictCount[currentDistrictName] += 1;

                                }
                            }
                        }
                        #endregion
                        sessionRow["Session A Count"] = currentSessionACount;
                        sessionRow["Session B Count"] = currentSessionBCount;
                        sessionRow["Session C Count"] = currentSessionCCount;
                    }
                    #endregion

                }
            }
            #region Somehow this person missed a schedule
            foreach (DataRow currentStudentRow in studentTable.Rows)
            {
                if (((String)currentStudentRow["Student Session A"]).Equals("0"))
                {
                    foreach (KeyValuePair<String, int> currentClass in PopularClasses)
                    {
                        String searchValue = "[SessionPresenter] = " + currentClass.Key;
                        DataRow[] sessionRows = sessionTable.Select(searchValue);
                        if (sessionRows.Count() == 0)
                        {
                            continue;
                        }
                        DataRow sessionRow = sessionRows[0];
                        DataRow roomRow = roomTable.Rows.Find(sessionRow["Session Room"]);
                        int currentSessionCount = (Int32)sessionRow["Session A Count"];
                        int presenterRoomCapacity = (Int32)roomRow["Room Capacity"];

                        if (currentSessionCount < presenterRoomCapacity)
                        {
                            currentStudentRow["Student Session A"] = currentClass.Key;
                            sessionRow["Session A Count"] = ++currentSessionCount;
                        }
                    }
                }
                if (((String)currentStudentRow["Student Session B"]).Equals("0"))
                {
                    foreach (KeyValuePair<String, int> currentClass in PopularClasses)
                    {
                        String searchValue = "[SessionPresenter] = " + currentClass.Key;
                        DataRow[] sessionRows = sessionTable.Select(searchValue);
                        if (sessionRows.Count() == 0)
                        {
                            continue;
                        }
                        DataRow sessionRow = sessionRows[0];
                        DataRow roomRow = roomTable.Rows.Find(sessionRow["Session Room"]);
                        int currentSessionCount = (Int32)sessionRow["Session B Count"];
                        int presenterRoomCapacity = (Int32)roomRow["Room Capacity"];

                        if (currentSessionCount < presenterRoomCapacity)
                        {
                            currentStudentRow["Student Session B"] = currentClass.Key;
                            sessionRow["Session B Count"] = ++currentSessionCount;
                        }
                    }
                }
                if (((String)currentStudentRow["Student Session C"]).Equals("0"))
                {
                    foreach (KeyValuePair<String, int> currentClass in PopularClasses)
                    {
                        String searchValue = "[SessionPresenter] = " + currentClass.Key;
                        DataRow[] sessionRows = sessionTable.Select(searchValue);
                        if (sessionRows.Count() == 0)
                        {
                            continue;
                        }
                        DataRow sessionRow = sessionRows[0];
                        DataRow roomRow = roomTable.Rows.Find(sessionRow["Session Room"]);
                        int currentSessionCount = (Int32)sessionRow["Session C Count"];
                        int presenterRoomCapacity = (Int32)roomRow["Room Capacity"];

                        if (currentSessionCount < presenterRoomCapacity)
                        {
                            currentStudentRow["Student Session C"] = currentClass.Key;
                            sessionRow["Session C Count"] = ++currentSessionCount;
                        }
                    }
                }
            }
            #endregion
        }


        private void checkDistrictPercentage()
        {
            int totalPercentage = 0;
            foreach (DataRow currentRow in districtTable.Rows)
            {
                totalPercentage += (Int32)currentRow["DistrictPercentage"];
            }

            if (!(totalPercentage == 100))
            {
                statusText = "District Percentages do not total 100 \nResubmit the file";
            }

        }


        private void createWorkshopStudentExcel()
        {
            FileInfo currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Workshop Final Schedule.xlsx");
            if (currentFile.Exists)
            {
                currentFile.Delete();  // ensures we create a new workbook
                currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Workshop Final Schedule.xlsx");
            }


            using (ExcelPackage currentExcel = new ExcelPackage(currentFile))
            {
                ExcelWorksheet currentSheet = currentExcel.Workbook.Worksheets.Add("Student Registrations");

                currentSheet.Cells["A1"].LoadFromDataTable(sessionTable, true);
                currentSheet.Cells["A1"].AutoFitColumns();
                String cellRange = "A1:" + Convert.ToChar('A' + sessionTable.Columns.Count - 1) + 1;

                using (ExcelRange currentRange = currentSheet.Cells[cellRange])
                {
                    currentRange.Style.WrapText = false;
                    currentRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    currentRange.Style.Font.Bold = true;
                    currentRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    currentRange.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    currentRange.Style.Font.Color.SetColor(Color.White);

                }

                String rowsCellRange = "A2:" + Convert.ToChar('A' + sessionTable.Columns.Count - 1) + sessionTable.Rows.Count * sessionTable.Columns.Count;

                using (ExcelRange currentRange = currentSheet.Cells[rowsCellRange])
                {
                    currentRange.Style.WrapText = false;
                    currentRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                }

                String currentPresenterTitle, currentStudentFirstName, currentStudentLastName, currentSessionA, currentSessionB, currentSessionC;
                int currentRowCounterA = 2, currentRowCounterB = 2, currentRowCounterC = 2;


                foreach (DataRow currentPresenterRow in presenterTable.Rows)
                {
                    currentPresenterTitle = (String)currentPresenterRow["Presenter Title"];

                    ExcelWorksheet currentSheetA = currentExcel.Workbook.Worksheets.Add(currentPresenterTitle + " A");
                    ExcelWorksheet currentSheetB = currentExcel.Workbook.Worksheets.Add(currentPresenterTitle + " B");
                    ExcelWorksheet currentSheetC = currentExcel.Workbook.Worksheets.Add(currentPresenterTitle + " C");

                    currentSheetA.Cells["A1"].Value = currentPresenterTitle;
                    currentSheetA.Cells["B1"].Value = "Session A";

                    currentSheetB.Cells["A1"].Value = currentPresenterTitle;
                    currentSheetB.Cells["B1"].Value = "Session B";

                    currentSheetC.Cells["A1"].Value = currentPresenterTitle;
                    currentSheetC.Cells["B1"].Value = "Session C";


                    foreach (DataRow currentStudentRow in sessionTable.Rows)
                    {
                        currentStudentFirstName = (String)currentStudentRow["Student First Name"];
                        currentStudentLastName = (String)currentStudentRow["Student Last Name"];
                        currentSessionA = (String)currentStudentRow["Student Session 1"];
                        currentSessionB = (String)currentStudentRow["Student Session 2"];
                        currentSessionC = (String)currentStudentRow["Student Session 3"];

                        if (currentSessionA.Equals(currentPresenterTitle))
                        {
                            //write student name to currentsheetA
                            currentSheetA.Cells[currentRowCounterA, 1].Value = currentStudentFirstName;
                            currentSheetA.Cells[currentRowCounterA, 2].Value = currentStudentLastName;
                            currentRowCounterA++;
                        }
                        if (currentSessionB.Equals(currentPresenterTitle))
                        {
                            //write student name to currentsheetB
                            currentSheetB.Cells[currentRowCounterB, 1].Value = currentStudentFirstName;
                            currentSheetB.Cells[currentRowCounterB, 2].Value = currentStudentLastName;
                            currentRowCounterB++;
                        }
                        if (currentSessionC.Equals(currentPresenterTitle))
                        {
                            //write student name to currentsheetC
                            currentSheetC.Cells[currentRowCounterC, 1].Value = currentStudentFirstName;
                            currentSheetC.Cells[currentRowCounterC, 2].Value = currentStudentLastName;
                            currentRowCounterC++;
                        }


                    }

                    currentRowCounterA = 2;
                    currentRowCounterB = 2;
                    currentRowCounterC = 2;
                }

                currentExcel.Save();
            }


        }

        #endregion
    }
}
