using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Data;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;

namespace WFC_Scheduler.Model
{
    class WorkshopData
    {


        private int canyonsPercent, granitePercent, murrayPercent, jordanPercent, saltLakePercent, tooelePercent, districtTotals;

        private string requestFilePath, workshopInfoFilePath, statusText, failedCutoff, exceedsCount, exceedsMessage;
        private DateTime cutoffTime;
        //private DataTable roomTable, schoolTable, districtTable, presenterTable, requestTable, studentTable, sessionTable;
        private Dictionary<String, Int32> PopularClasses, FourthClasses, ThirdClasses, SecondClasses, FirstClasses;
        private List<DistrictClass> districtClassCount;
        private List<Session> stillOneBlank, stillTwoBlank, stillThreeBlank;
        private List<DistrictClass> stupidList;

        private bool okToSchedule;

        private List<Room> roomList;
        private List<School> schoolList;
        private List<District> districtList;
        private List<Presenter> presenterList;
        private List<Request> requestList;
        private List<Student> studentList;
        private List<Session> sessionList;

        public List<District> DistrictList
        {
            get { return districtList; }
            set { districtList = value; }
        }

        public bool ScheduleOK
        {
            get { return okToSchedule; }
            set { okToSchedule = value; }
        }

        public DateTime CutoffTime
        {
            get { return cutoffTime; }
            set { cutoffTime = value; }
        }

        public int CanyonsPercent
        {
            get { return canyonsPercent; }
            set { canyonsPercent = value; }
        }

        public int GranitePercent
        {
            get { return granitePercent; }
            set { granitePercent = value; }
        }

        public int MurrayPercent
        {
            get { return murrayPercent; }
            set { murrayPercent = value; }
        }

        public int JordanPercent
        {
            get { return jordanPercent; }
            set { jordanPercent = value; }
        }

        public int SaltLakePercent
        {
            get { return saltLakePercent; }
            set { saltLakePercent = value; }
        }

        public int TooelePercent
        {
            get { return tooelePercent; }
            set { tooelePercent = value; }
        }

        public string RequestFilePath
        {
            get { return requestFilePath; }
            set { requestFilePath = value; }
        }

        public string WorkshopInfoFilePath
        {
            get { return workshopInfoFilePath; }
            set { workshopInfoFilePath = value; }
        }

        public WorkshopData()
        {
            canyonsPercent = 0;
            granitePercent = 0;
            murrayPercent = 0;
            jordanPercent = 0;
            saltLakePercent = 0;
            tooelePercent = 0;
            districtTotals = 0;
            okToSchedule = false;
            statusText = "Creating Workshop data: ";
            failedCutoff = "No sessions, enrolled after cutoff time";
            exceedsCount = "District has exceeded registration capability, student not scheduled";
        }

        public int DistrictTotals
        {
            get { return canyonsPercent + granitePercent + murrayPercent + jordanPercent + saltLakePercent + tooelePercent; }
        }

        public string StatusText
        {
            get { return statusText; }
        }

        /// <summary>
        /// Determine popularity of classes for the selection process using the lists
        /// </summary>
        public void verifyPopularity()
        {
            PopularClasses = new Dictionary<String, int>();
            FourthClasses = new Dictionary<String, int>();
            ThirdClasses = new Dictionary<String, int>();
            SecondClasses = new Dictionary<String, int>();
            FirstClasses = new Dictionary<String, int>();

            foreach (Presenter currentPresenter in presenterList)
            {
                PopularClasses.Add(currentPresenter.PresenterTitle, 0);

                #region Loop over all requests and populate dictionaries
                foreach (Request currentRequest in requestList)
                {
                    string presenterName = currentPresenter.PresenterTitle;
                    if (currentRequest.RequestOne.Equals(presenterName) || currentRequest.RequestTwo.Equals(presenterName) || currentRequest.RequestThree.Equals(presenterName) || currentRequest.RequestFour.Equals(presenterName) || currentRequest.RequestFive.Equals(presenterName))
                    {
                        if (PopularClasses.ContainsKey(presenterName))
                        {
                            int currentCount = PopularClasses[presenterName];
                            ++currentCount;
                            PopularClasses[presenterName] = currentCount;
                        }
                        else
                        {
                            PopularClasses.Add(presenterName, 1);
                        }
                    }

                    if (currentRequest.RequestOne.Equals(presenterName) || currentRequest.RequestTwo.Equals(presenterName) || currentRequest.RequestThree.Equals(presenterName) || currentRequest.RequestFour.Equals(presenterName))
                    {
                        if (FourthClasses.ContainsKey(presenterName))
                        {
                            int currentCount = FourthClasses[presenterName];
                            ++currentCount;
                            FourthClasses[presenterName] = currentCount;
                        }
                        else
                        {
                            FourthClasses.Add(presenterName, 1);
                        }
                    }

                    if (currentRequest.RequestOne.Equals(presenterName) || currentRequest.RequestTwo.Equals(presenterName) || currentRequest.RequestThree.Equals(presenterName))
                    {
                        if (ThirdClasses.ContainsKey(presenterName))
                        {
                            int currentCount = ThirdClasses[presenterName];
                            ++currentCount;
                            ThirdClasses[presenterName] = currentCount;
                        }
                        else
                        {
                            ThirdClasses.Add(presenterName, 1);
                        }
                    }

                    if (currentRequest.RequestOne.Equals(presenterName) || currentRequest.RequestTwo.Equals(presenterName))
                    {
                        if (SecondClasses.ContainsKey(presenterName))
                        {
                            int currentCount = SecondClasses[presenterName];
                            ++currentCount;
                            SecondClasses[presenterName] = currentCount;
                        }
                        else
                        {
                            SecondClasses.Add(presenterName, 1);
                        }
                    }

                    if (currentRequest.RequestOne.Equals(presenterName))
                    {
                        if (FirstClasses.ContainsKey(presenterName))
                        {
                            int currentCount = FirstClasses[presenterName];
                            ++currentCount;
                            FirstClasses[presenterName] = currentCount;
                        }
                        else
                        {
                            FirstClasses.Add(presenterName, 1);
                        }
                    }
                }
                #endregion
            }

            PopularClasses = PopularClasses.OrderByDescending(x => x.Value).ToDictionary(x => x.Key, x => x.Value);
        }

        /// <summary>
        /// Checker method for supplied information.
        /// </summary>
        private void checkDistrictPercentage()
        {
            int totalPercentage = 0;
            foreach (District currentDistrict in districtList)
            {
                totalPercentage += currentDistrict.DistrictPercentage;
            }

            if (!(totalPercentage == 100))
            {
                statusText = "District Percentages do not total 100 \nResubmit the file";
                okToSchedule = false;
            }

        }

        /// <summary>
        /// Checker method for correct data sheets in submitted XLSX file.
        /// </summary>
        /// <param name="current">The ExcelWorkbook file</param>
        /// <returns>Boolean value based on whether all needed sheets are present.</returns>
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

        /// <summary>
        /// Method to bring all data from the XLSX file to associated lists.
        /// </summary>
        private void importExcelDataToLists()
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

                        createRoomList(roomSheet);
                        createPresenterList(presentersSheet);
                        createDistrictList(districtSheet);
                        checkDistrictPercentage();
                        createSchoolList(schoolsSheet);
                        createSessionList();

                    }
                }
            }
            statusText += "\nWorkshop data extracted: \n     room, presenter, district, school, and session lists created.\n";


            FileInfo requestFile = new FileInfo(requestFilePath);
            using (ExcelPackage currentExcelFile = new ExcelPackage(requestFile))
            {
                ExcelWorkbook currentWorkbook = currentExcelFile.Workbook;
                if (currentWorkbook != null)
                {
                    ExcelWorksheet requestSheet = currentWorkbook.Worksheets[1];
                    createStudentAndRequestList(requestSheet);
                }

            }
            statusText += "Student requests extracted:\n     student and request lists generated";

        }

        #region Create Lists

        /// <summary>
        /// Creates a list of all district information.
        /// </summary>
        /// <param name="districtSheet"></param>
        private void createDistrictList(ExcelWorksheet districtSheet)
        {
            if (districtList != null && districtList.Count != 0)
            {
                statusText = "Workshop data recreated with updated district percentages";
            }

            else if (canyonsPercent != 0 || granitePercent != 0 || tooelePercent != 0 || murrayPercent != 0 || jordanPercent != 0 || saltLakePercent != 0)
            {
                districtList = new List<District>();
                //fix to import next .xlsx.
                districtList.Add(new District("Canyons", canyonsPercent, 93));
                districtList.Add(new District("Granite", granitePercent, 160));
                districtList.Add(new District("Jordan", jordanPercent, 121));
                districtList.Add(new District("Murray", murrayPercent, 20));
                districtList.Add(new District("Tooele", tooelePercent, 59));
                districtList.Add(new District("Salt Lake City", saltLakePercent, 35));

                statusText = "Workshop data recreated with updated district percentages";
            }
            else
            {
                districtList = new List<District>();

                String districtName;
                int districtPercentage, districtMax;
                //int[] stuffme = { 93, 160, 121, 20, 59, 35 };
                for (int row = 2; row <= districtSheet.Dimension.End.Row; row++)
                {
                    districtName = (String)districtSheet.Cells[row, 1].Value;
                    districtPercentage = Convert.ToInt32(districtSheet.Cells[row, 2].Value);
                    districtMax = Convert.ToInt32(districtSheet.Cells[row, 3].Value);

                    districtList.Add(new District(districtName, districtPercentage, districtMax));
                }
            }

        }

        /// <summary>
        /// Creates the school list, must be completed after the district list is created as it does a search on that data.
        /// </summary>
        /// <param name="schoolsSheet">The extracted Excel 2007+ data</param>
        private void createSchoolList(ExcelWorksheet schoolsSheet)
        {
            schoolList = new List<School>();

            String schoolName, districtName, errorSchool, errorDistrict;
            errorSchool = "";
            errorDistrict = "";
            try
            {
                for (int row = 2; row <= schoolsSheet.Dimension.End.Row; row++)
                {
                    schoolName = (String)schoolsSheet.Cells[row, 1].Value;
                    districtName = (String)schoolsSheet.Cells[row, 2].Value;

                    District currentDistrict = districtList.Find(delegate(District current) { return current.DistrictName.Equals(districtName); });

                    schoolList.Add(new School(currentDistrict, schoolName));
                    errorSchool = schoolName;
                    errorDistrict = districtName;
                }
            }
            catch (Exception currentException)
            {
                MessageBox.Show("There was an error in the schools.  Make sure that each school is associated with a district");
                MessageBox.Show("The last successful school was: " + errorSchool + " check the spelling of the districts in the data file");

            }
        }

        /// <summary>
        /// Creates the room list
        /// </summary>
        /// <param name="roomSheet">Extracted Excel 2007+ data</param>
        private void createRoomList(ExcelWorksheet roomSheet)
        {
            roomList = new List<Room>();
            String roomName, errorRoom;
            int roomCapacity;
            errorRoom = "";

            try
            {
                for (int row = 2; row <= roomSheet.Dimension.End.Row; row++)
                {
                    roomName = (String)roomSheet.Cells[row, 1].Value.ToString();
                    roomCapacity = Convert.ToInt32(roomSheet.Cells[row, 2].Value);
                    roomList.Add(new Room(roomName, roomCapacity));
                    errorRoom = roomName;
                }
            }
            catch (Exception currentException)
            {
                MessageBox.Show("There was an error when reading in the rooms from the data file.");
                MessageBox.Show("The last successful room was: " + errorRoom + " check the room names are all different in the data file");

            }




        }

        /// <summary>
        /// Creates the presenter list.  Must be completed after the room list as it uses that for a search for specific room object
        /// </summary>
        /// <param name="presentersSheet"></param>
        private void createPresenterList(ExcelWorksheet presentersSheet)
        {
            presenterList = new List<Presenter>();

            String PresenterTitle, PresenterDescription, presenterRoom, errorPresenter;
            errorPresenter = "";
            try
            {
                for (int row = 2; row <= presentersSheet.Dimension.End.Row; row++)
                {
                    PresenterTitle = (String)presentersSheet.Cells[row, 1].Value;
                    PresenterDescription = (String)presentersSheet.Cells[row, 2].Value;
                    presenterRoom = (String)presentersSheet.Cells[row, 3].Value.ToString();

                    Room currentRoom = roomList.Find(delegate(Room current) { return current.RoomName.Equals(presenterRoom); });
                    presenterList.Add(new Presenter(PresenterTitle, PresenterDescription, currentRoom));
                    errorPresenter = PresenterTitle;
                }
            }
            catch (Exception currentException)
            {
                MessageBox.Show("There was an error reading the presenters.  Make sure that all the fields are present");
                MessageBox.Show("last successful presenter was: " + errorPresenter + " check the data file");
            }
        }

        /// <summary>
        /// Creates the linked student and request sheets
        /// </summary>
        /// <param name="requestSheet">Extracted Excel 2007+ data</param>
        private void createStudentAndRequestList(ExcelWorksheet requestSheet)
        {

            string firstName, lastName, email, phone, requestOne, requestTwo, requestThree, requestFour, requestFive, school, studentGradeLevel, studentTeacher, studentNumber, errorStudent;
            int studentID;
            DateTime studentTime;
            errorStudent = "";
            studentList = new List<Student>();
            requestList = new List<Request>();

            try
            {
                for (int row = 2; row <= requestSheet.Dimension.End.Row; row++)
                {
                    studentID = row - 1;
                    String test = requestSheet.Cells[row, 1].Value.ToString();
                    try
                    {
                        studentTime = DateTime.FromOADate(Double.Parse(test));

                    }
                    catch (Exception parseException)
                    {
                        DateTime.TryParse(test, out studentTime);
                    }

                    DateTime.TryParse(test, out studentTime);
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

                    errorStudent = lastName + ", " + firstName;

                    bool ok = false;
                    School currentSchool = schoolList.Find(delegate(School temp) { return temp.SchoolName.Equals(school); });
                    if (currentSchool != null)
                    {
                        ok = true;
                    }
                    District currentDistrict = currentSchool.SchoolDistrict;
                    if (currentDistrict.TotalStudents < currentDistrict.MaxStudents)
                    {
                        currentDistrict.TotalStudents++;
                        ok = true;
                    }
                    else
                    {
                        currentDistrict.TotalStudents++;
                        ok = false;
                    }

                    Request currentRequest;
                    Student currentStudent = new Student(firstName, lastName, studentNumber, email, studentTeacher, phone, studentGradeLevel, school, ok);
                    currentRequest = new Request(currentStudent, studentTime, requestOne, requestTwo, requestThree, requestFour, requestFive);
                    currentStudent.StudentRequest = currentRequest;

                    requestList.Add(currentRequest);
                    studentList.Add(currentStudent);

                }

            }
            catch (Exception currentException)
            {
                MessageBox.Show("There was an error reading in the excel file.  The problem occurred in the student import");
                MessageBox.Show("The last successful student was: " + errorStudent + " check that the school names match between the datafile and the request file");

            }
        }

        /// <summary>
        /// Creates the session list.  Must occur after the presenter list is created.
        /// </summary>
        private void createSessionList()
        {
            sessionList = new List<Session>();
            foreach (Presenter currentPresenter in presenterList)
            {
                sessionList.Add(new Session(currentPresenter));
            }

        }

        #endregion

        /// <summary>
        /// Main method for generating schedules for the Wasatch Front Consortium project.
        /// </summary>
        public void generateScheduleFromLists()
        {
            List<Student> blankStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionOne.Length == 0 || tempStudent.SessionTwo.Length == 0 || tempStudent.SessionThree.Length == 0 || tempStudent.SessionOne.Equals(exceedsCount) || tempStudent.SessionTwo.Equals(exceedsCount) || tempStudent.SessionThree.Equals(exceedsCount)); });

            List<Student> exceedsStudents = studentList.FindAll(delegate(Student exceeds) { return (!exceeds.OkToSchedule); });
            foreach (Student exceedingStudent in exceedsStudents)
            {
                exceedingStudent.SessionOne = exceedsCount;
                exceedingStudent.SessionTwo = exceedsCount;
                exceedingStudent.SessionThree = exceedsCount;
            }
            exceedsMessage = "\n" + exceedsStudents.Count + " students exceeded district counts\n";
            statusText += exceedsMessage;
            districtClassCount = new List<DistrictClass>();

            foreach (District currentDistrict in districtList)
            {
                foreach (Presenter currentPresenter in presenterList)
                {
                    districtClassCount.Add(new DistrictClass(currentDistrict.DistrictName, currentPresenter.PresenterTitle));
                }
            }



            #region Basic scheduling
            foreach (KeyValuePair<string, int> classNameAndCount in PopularClasses)
            {
                Session currentSession = sessionList.Find(delegate(Session current) { return current.SessionPresenter.PresenterTitle.Equals(classNameAndCount.Key); });
                Dictionary<String, Int32> currentDistrictCount = new Dictionary<string, int>();



                Room currentRoom = roomList.Find(delegate(Room tempRoom) { return tempRoom.Equals(currentSession.SessionPresenter.PresenterRoom); });
                int totalRoomCapacity = 3 * currentRoom.RoomCapacity;
                int currentCapacity = currentRoom.RoomCapacity;

                int currentSessionACount = 0;
                int currentSessionBCount = 0;
                int currentSessionCCount = 0;

                List<Request> wantedClasses = requestList.FindAll(delegate(Request tempRequest) { return (tempRequest.RequestFive.Equals(currentSession.SessionPresenter.PresenterTitle) || tempRequest.RequestFour.Equals(currentSession.SessionPresenter.PresenterTitle) || tempRequest.RequestThree.Equals(currentSession.SessionPresenter.PresenterTitle) || tempRequest.RequestTwo.Equals(currentSession.SessionPresenter.PresenterTitle) || tempRequest.RequestOne.Equals(currentSession.SessionPresenter.PresenterTitle)); });

                int smallTest = 0;

                foreach (Request currentRequest in wantedClasses)
                {
                    if (currentRequest.RequestTime > cutoffTime)
                    {

                        Student tempStudent = studentList.Find(delegate(Student temp) { return temp.Equals(currentRequest.RequestingStudent); });
                        tempStudent.SessionOne = failedCutoff;
                        tempStudent.SessionTwo = failedCutoff;
                        tempStudent.SessionThree = failedCutoff;

                    }

                    else
                    {
                        Student currentStudent = studentList.Find(delegate(Student temp) { return temp.Equals(currentRequest.RequestingStudent); });
                        School currentSchool = schoolList.Find(delegate(School temp) { return temp.SchoolName.Equals(currentStudent.SchoolName); });
                        District currentDistrict = currentSchool.SchoolDistrict;



                        if (currentDistrict.DistrictName.Equals("Murray"))
                        {
                            smallTest++;
                        }


                        double currentPercentage = (((double)currentDistrict.DistrictPercentage) / 100.00);
                        int currentDistrictMax = (Int32)Math.Ceiling((currentPercentage * (double)totalRoomCapacity));
                        DistrictClass currentClassCount = districtClassCount.Find(delegate(DistrictClass current) { return (current.DistrictName.Equals(currentDistrict.DistrictName) && current.ClassName.Equals(classNameAndCount.Key)); });
                        currentClassCount.Max = currentDistrictMax;
                        //if(currentStudent.SessionOne.Length != 0)

                        #region All requests will fit
                        if (classNameAndCount.Value < totalRoomCapacity)
                        {
                            if (currentRequest.RequestOne.Equals(classNameAndCount.Key) || currentRequest.RequestTwo.Equals(classNameAndCount.Key) || currentRequest.RequestThree.Equals(classNameAndCount.Key) || currentRequest.RequestFour.Equals(classNameAndCount.Key) || currentRequest.RequestFive.Equals(classNameAndCount.Key))
                            {
                                if (currentClassCount.Count < currentDistrictMax && currentStudent.OkToSchedule)
                                {

                                    int randomizer = (currentSessionACount + currentSessionBCount + currentSessionCCount) % 3;
                                    if (randomizer == 0)
                                    {
                                        if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                    else if (randomizer == 1)
                                    {
                                        if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                    }
                                    else
                                    {
                                        if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }

                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        #region Fourth and better requests
                        else if (!FourthClasses.ContainsKey(classNameAndCount.Key))
                        {
                            continue;
                        }
                        else if (FourthClasses[classNameAndCount.Key] < totalRoomCapacity && FourthClasses.ContainsKey(classNameAndCount.Key))
                        {
                            if (currentRequest.RequestOne.Equals(classNameAndCount.Key) || currentRequest.RequestTwo.Equals(classNameAndCount.Key) || currentRequest.RequestThree.Equals(classNameAndCount.Key) || currentRequest.RequestFour.Equals(classNameAndCount.Key))
                            {
                                if (currentClassCount.Count < currentDistrictMax && currentStudent.OkToSchedule)
                                {

                                    int randomizer = (currentSessionACount + currentSessionBCount + currentSessionCCount) % 3;
                                    if (randomizer == 0)
                                    {

                                        if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                    else if (randomizer == 1)
                                    {
                                        if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                    }
                                    else
                                    {
                                        if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }

                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        #region Third and better requests
                        else if (!ThirdClasses.ContainsKey(classNameAndCount.Key))
                        {
                            continue;
                        }
                        else if (ThirdClasses[classNameAndCount.Key] < totalRoomCapacity && ThirdClasses.ContainsKey(classNameAndCount.Key))
                        {
                            if (currentRequest.RequestOne.Equals(classNameAndCount.Key) || currentRequest.RequestTwo.Equals(classNameAndCount.Key) || currentRequest.RequestThree.Equals(classNameAndCount.Key))
                            {
                                if (currentClassCount.Count < currentDistrictMax && currentStudent.OkToSchedule)
                                {

                                    int randomizer = (currentSessionACount + currentSessionBCount + currentSessionCCount) % 3;
                                    if (randomizer == 0)
                                    {
                                        if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                    else if (randomizer == 1)
                                    {
                                        if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }

                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                    }
                                    else
                                    {
                                        if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        #region Second and better requests
                        else if (!SecondClasses.ContainsKey(classNameAndCount.Key))
                        {
                            continue;
                        }
                        else if (SecondClasses[classNameAndCount.Key] < totalRoomCapacity && SecondClasses.ContainsKey(classNameAndCount.Key))
                        {
                            if (currentRequest.RequestOne.Equals(classNameAndCount.Key) || currentRequest.RequestTwo.Equals(classNameAndCount.Key))
                            {
                                if (currentClassCount.Count < currentDistrictMax && currentStudent.OkToSchedule)
                                {
                                    int randomizer = (currentSessionACount + currentSessionBCount + currentSessionCCount) % 3;
                                    if (randomizer == 0)
                                    {
                                        if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                    else if (randomizer == 1)
                                    {
                                        if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                    }
                                    else
                                    {
                                        if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                }
                            }

                        }
                        #endregion

                        #region First only requests
                        else if (!FirstClasses.ContainsKey(classNameAndCount.Key))
                        {
                            continue;
                        }
                        else if (FirstClasses[classNameAndCount.Key] < totalRoomCapacity && FirstClasses.ContainsKey(classNameAndCount.Key))
                        {
                            if (currentRequest.RequestOne.Equals(classNameAndCount.Key))
                            {
                                if (currentClassCount.Count < currentDistrictMax && currentStudent.OkToSchedule)
                                {

                                    int randomizer = (currentSessionACount + currentSessionBCount + currentSessionCCount) % 3;
                                    if (randomizer == 0)
                                    {
                                        if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                    else if (randomizer == 1)
                                    {
                                        if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                    }
                                    else
                                    {
                                        if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        #region Less than first choice requests

                        else
                        {
                            if (currentRequest.RequestOne.Equals(classNameAndCount.Key))
                            {
                                if (currentClassCount.Count < currentDistrictMax && currentStudent.OkToSchedule)
                                {

                                    int randomizer = (currentSessionACount + currentSessionBCount + currentSessionCCount) % 3;
                                    if (randomizer == 0)
                                    {
                                        if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                    else if (randomizer == 1)
                                    {
                                        if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                    }
                                    else
                                    {
                                        if ((currentStudent.SessionOne.Length == 0) && (currentSessionACount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionOne = classNameAndCount.Key;
                                            currentSessionACount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionOneList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionTwo.Length == 0) && (currentSessionBCount < currentCapacity) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionTwo = classNameAndCount.Key;
                                            currentSessionBCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionTwoList.Add(currentStudent);
                                        }
                                        else if ((currentStudent.SessionThree.Length == 0) && (currentSessionCCount < currentCapacity) && !(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle) && !(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle))))
                                        {
                                            currentStudent.SessionThree = classNameAndCount.Key;
                                            currentSessionCCount++;
                                            currentClassCount.Count++;
                                            currentSession.SessionThreeList.Add(currentStudent);
                                        }
                                    }
                                }
                            }
                        }
                        #endregion


                    }


                }


            }
            #endregion

            extraScheduling();

            blankStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionOne.Length == 0 || tempStudent.SessionTwo.Length == 0 || tempStudent.SessionThree.Length == 0); });

            List<DistrictClass> otherList = districtClassCount.FindAll(delegate(DistrictClass tempClassCount) { return (tempClassCount.Count < tempClassCount.Max); });


            foreach (Student currentStudent in blankStudents)
            {
                clearSchedule(currentStudent, true);

            }

            checkDuplicates();
            statusText += exceedsMessage + "\n" + blankStudents.Count + " students were not scheduled: Counts exceeded";


        }

        /// <summary>
        /// Continuing the scheduling to get all straggling schedules.
        /// </summary>
        private void extraScheduling()
        {
            #region Somehow scheduling missed this student


            List<Student> emptyOneStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionOne.Length == 0); });
            foreach (Student missingFirst in emptyOneStudents)
            {
                stillOneBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionOneCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });

                stillOneBlank = stillOneBlank.OrderBy(x => x.SessionOneCount).ToList<Session>();

                foreach (Session emptyOne in stillOneBlank)
                {
                    if (missingFirst.SessionOne.Length == 0)
                    {
                        int currentCount = emptyOne.SessionOneCount;
                        int currentCap = emptyOne.Capacity;
                        School currentSchool = schoolList.Find(delegate(School temp) { return temp.SchoolName.Equals(missingFirst.SchoolName); });
                        District currentDistrict = currentSchool.SchoolDistrict;

                        DistrictClass currentClassCount = districtClassCount.Find(delegate(DistrictClass current) { return (current.DistrictName.Equals(currentDistrict.DistrictName) && current.ClassName.Equals(emptyOne.SessionPresenter.PresenterTitle)); });

                        if (currentCount < currentCap && missingFirst.OkToSchedule && !missingFirst.SessionTwo.Equals(emptyOne.SessionPresenter.PresenterTitle) && !missingFirst.SessionThree.Equals(emptyOne.SessionPresenter.PresenterTitle))
                        {
                            if (currentClassCount.Count < currentClassCount.Max)
                            {
                                missingFirst.SessionOne = emptyOne.SessionPresenter.PresenterTitle;
                                currentClassCount.Count++;
                                emptyOne.SessionOneList.Add(missingFirst);
                            }
                        }
                    }
                }
            }
            emptyOneStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionOne.Length == 0); });
            stillOneBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionOneCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });
            int testOne;

            List<Student> emptyTwoStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionTwo.Length == 0); });
            foreach (Student missingSecond in emptyTwoStudents)
            {
                stillTwoBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionTwoCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });

                stillTwoBlank = stillTwoBlank.OrderBy(x => x.SessionTwoCount).ToList<Session>();

                foreach (Session emptyTwo in stillTwoBlank)
                {
                    if (missingSecond.SessionTwo.Length == 0)
                    {
                        int currentCount = emptyTwo.SessionTwoCount;
                        int currentCap = emptyTwo.Capacity;
                        School currentSchool = schoolList.Find(delegate(School temp) { return temp.SchoolName.Equals(missingSecond.SchoolName); });
                        District currentDistrict = currentSchool.SchoolDistrict;

                        DistrictClass currentClassCount = districtClassCount.Find(delegate(DistrictClass current) { return (current.DistrictName.Equals(currentDistrict.DistrictName) && current.ClassName.Equals(emptyTwo.SessionPresenter.PresenterTitle)); });

                        if (currentCount < currentCap && missingSecond.OkToSchedule && !missingSecond.SessionOne.Equals(emptyTwo.SessionPresenter.PresenterTitle) && !missingSecond.SessionThree.Equals(emptyTwo.SessionPresenter.PresenterTitle))
                        {
                            if (currentClassCount.Count < currentClassCount.Max)
                            {
                                missingSecond.SessionTwo = emptyTwo.SessionPresenter.PresenterTitle;
                                currentClassCount.Count++;
                                emptyTwo.SessionTwoList.Add(missingSecond);
                            }

                        }
                    }
                }
            }
            emptyTwoStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionTwo.Length == 0); });
            stillTwoBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionTwoCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });
            int testTwo;


            List<Student> emptyThreeStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionThree.Length == 0); });
            foreach (Student missingThird in emptyThreeStudents)
            {
                stillThreeBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionThreeCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });

                stillThreeBlank = stillThreeBlank.OrderBy(x => x.SessionThreeCount).ToList<Session>();

                foreach (Session emptyThree in stillThreeBlank)
                {
                    if (missingThird.SessionThree.Length == 0)
                    {

                        int currentCount = emptyThree.SessionThreeCount;
                        int currentCap = emptyThree.Capacity;
                        School currentSchool = schoolList.Find(delegate(School temp) { return temp.SchoolName.Equals(missingThird.SchoolName); });
                        District currentDistrict = currentSchool.SchoolDistrict;

                        DistrictClass currentClassCount = districtClassCount.Find(delegate(DistrictClass current) { return (current.DistrictName.Equals(currentDistrict.DistrictName) && current.ClassName.Equals(emptyThree.SessionPresenter.PresenterTitle)); });

                        if (currentCount < currentCap && missingThird.OkToSchedule && !missingThird.SessionTwo.Equals(emptyThree.SessionPresenter.PresenterTitle) && !missingThird.SessionOne.Equals(emptyThree.SessionPresenter.PresenterTitle))
                        {
                            if (currentClassCount.Count < currentClassCount.Max)
                            {
                                missingThird.SessionThree = emptyThree.SessionPresenter.PresenterTitle;
                                currentClassCount.Count++;
                                emptyThree.SessionThreeList.Add(missingThird);
                            }
                        }
                    }
                }
            }
            emptyThreeStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionThree.Length == 0); });
            stillThreeBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionThreeCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });
            int testThree;

            List<Student> emptyStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionOne.Length == 0 || tempStudent.SessionTwo.Length == 0 || tempStudent.SessionThree.Length == 0); });

            foreach (Student currentStudent in emptyStudents)
            {
                stillOneBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionOneCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });
                stillTwoBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionTwoCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });
                stillThreeBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionThreeCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });

                stillOneBlank = stillOneBlank.OrderBy(x => x.SessionOneCount).ToList<Session>();
                stillTwoBlank = stillTwoBlank.OrderBy(x => x.SessionTwoCount).ToList<Session>();
                stillThreeBlank = stillThreeBlank.OrderBy(x => x.SessionThreeCount).ToList<Session>();

                if (stillOneBlank.Count > 0 && currentStudent.SessionOne.Length == 0)
                {
                    foreach (Session currentSession in stillOneBlank)
                    {
                        if (currentStudent.SessionOne.Length == 0)
                        {
                            int currentCount = currentSession.SessionOneCount;
                            int currentCap = currentSession.Capacity;
                            int totalCap = currentCap * 3;
                            School currentSchool = schoolList.Find(delegate(School temp) { return temp.SchoolName.Equals(currentStudent.SchoolName); });
                            District currentDistrict = currentSchool.SchoolDistrict;
                            double currentPercentage = (double)currentDistrict.DistrictPercentage / 100.00;
                            int currentDistrictMax = (Int32)Math.Ceiling(currentPercentage * (double)totalCap);
                            DistrictClass currentClassCount = districtClassCount.Find(delegate(DistrictClass current) { return (current.DistrictName.Equals(currentDistrict.DistrictName) && current.ClassName.Equals(currentSession.SessionPresenter.PresenterTitle)); });

                            if ((currentCount < currentCap) && currentStudent.OkToSchedule)
                            {
                                if (!(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle)))
                                {
                                    if (!(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle)))
                                    {
                                        currentStudent.SessionOne = currentSession.SessionPresenter.PresenterTitle;
                                        currentClassCount.Count++;
                                        currentSession.SessionOneList.Add(currentStudent);

                                    }
                                }

                            }
                        }

                    }
                }

                if (stillTwoBlank.Count > 0 && currentStudent.SessionTwo.Length == 0)
                {
                    foreach (Session currentSession in stillTwoBlank)
                    {
                        if (currentStudent.SessionTwo.Length == 0)
                        {
                            int currentCount = currentSession.SessionTwoCount;
                            int currentCap = currentSession.Capacity;
                            int totalCap = currentCap * 3;
                            School currentSchool = schoolList.Find(delegate(School temp) { return temp.SchoolName.Equals(currentStudent.SchoolName); });
                            District currentDistrict = currentSchool.SchoolDistrict;
                            double currentPercentage = (double)currentDistrict.DistrictPercentage / 100.00;
                            int currentDistrictMax = (Int32)Math.Ceiling(currentPercentage * (double)totalCap);
                            DistrictClass currentClassCount = districtClassCount.Find(delegate(DistrictClass current) { return (current.DistrictName.Equals(currentDistrict.DistrictName) && current.ClassName.Equals(currentSession.SessionPresenter.PresenterTitle)); });

                            if ((currentCount < currentCap) && currentStudent.OkToSchedule)
                            {
                                if (!(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle)))
                                {
                                    if (!(currentStudent.SessionThree.Equals(currentSession.SessionPresenter.PresenterTitle)))
                                    {
                                        currentStudent.SessionTwo = currentSession.SessionPresenter.PresenterTitle;
                                        currentClassCount.Count++;
                                        currentSession.SessionTwoList.Add(currentStudent);
                                    }

                                }

                            }
                        }
                    }

                }
                if (stillThreeBlank.Count > 0 && currentStudent.SessionThree.Length == 0)
                {
                    foreach (Session currentSession in stillThreeBlank)
                    {
                        if (currentStudent.SessionThree.Length == 0)
                        {
                            int currentCount = currentSession.SessionThreeCount;
                            int currentCap = currentSession.Capacity;
                            int totalCap = currentCap * 3;
                            School currentSchool = schoolList.Find(delegate(School temp) { return temp.SchoolName.Equals(currentStudent.SchoolName); });
                            District currentDistrict = currentSchool.SchoolDistrict;
                            double currentPercentage = (double)currentDistrict.DistrictPercentage / 100.00;
                            int currentDistrictMax = (Int32)Math.Ceiling(currentPercentage * (double)totalCap);
                            DistrictClass currentClassCount = districtClassCount.Find(delegate(DistrictClass current) { return (current.DistrictName.Equals(currentDistrict.DistrictName) && current.ClassName.Equals(currentSession.SessionPresenter.PresenterTitle)); });

                            if ((currentCount < currentCap) && currentStudent.OkToSchedule)
                            {
                                if (!(currentStudent.SessionTwo.Equals(currentSession.SessionPresenter.PresenterTitle)))
                                {
                                    if (!(currentStudent.SessionOne.Equals(currentSession.SessionPresenter.PresenterTitle)))
                                    {
                                        currentStudent.SessionThree = currentSession.SessionPresenter.PresenterTitle;
                                        currentClassCount.Count++;
                                        currentSession.SessionThreeList.Add(currentStudent);
                                    }
                                }

                            }
                        }

                    }
                }



            }

            emptyStudents = studentList.FindAll(delegate(Student tempStudent) { return (tempStudent.SessionOne.Length == 0 || tempStudent.SessionTwo.Length == 0 || tempStudent.SessionThree.Length == 0); });

            foreach (Student notScheduled in emptyStudents)
            {
                clearSchedule(notScheduled, false);
            }
            #endregion
            stupidList = districtClassCount.FindAll(delegate(DistrictClass tempClassCount) { return (tempClassCount.Count < tempClassCount.Max); });



        }

        /// <summary>
        /// Check for duplicates within the student schedules
        /// </summary>
        private void checkDuplicates()
        {
            int dupeCount = 0;
            List<Student> validStudents = studentList.FindAll(delegate(Student validStudent) { return (!validStudent.SessionOne.Equals(exceedsCount)); });
            foreach (Student duplicateStudent in validStudents)
            {
                if (duplicateStudent.SessionOne.Equals(duplicateStudent.SessionTwo))
                {
                    dupeCount++;
                }
                if (duplicateStudent.SessionThree.Equals(duplicateStudent.SessionTwo))
                {
                    dupeCount++;
                }
                if (duplicateStudent.SessionOne.Equals(duplicateStudent.SessionThree))
                {
                    dupeCount++;
                }
            }
            System.Windows.Forms.MessageBox.Show("Total duplicated classes: " + dupeCount);

        }

        /// <summary>
        /// Remove the current student's schedule.
        /// </summary>
        /// <param name="currentStudent"></param>
        /// <param name="outOfClass"></param>
        private void clearSchedule(Student currentStudent, Boolean outOfClass)
        {
            DistrictClass currentDistrictClassCount;
            foreach (Session current in sessionList)
            {
                School currentSchool = schoolList.Find(delegate(School currSchool) { return (currSchool.SchoolName.Equals(currentStudent.SchoolName)); });
                District currentDistrict = currentSchool.SchoolDistrict;
                currentDistrictClassCount = districtClassCount.Find(delegate(DistrictClass currentDC) { return (currentDC.ClassName.Equals(current.SessionPresenter.PresenterTitle) && currentDC.DistrictName.Equals(currentDistrict.DistrictName)); });

                if (current.SessionOneList.Contains(currentStudent) || current.SessionTwoList.Contains(currentStudent) || current.SessionThreeList.Contains(currentStudent))
                {
                    if (current.SessionOneList.Contains(currentStudent))
                    {
                        current.SessionOneList.Remove(currentStudent);
                        currentDistrictClassCount.Count--;
                    }
                    if (current.SessionTwoList.Contains(currentStudent))
                    {
                        current.SessionTwoList.Remove(currentStudent);
                        currentDistrictClassCount.Count--;
                    }
                    if (current.SessionThreeList.Contains(currentStudent))
                    {
                        current.SessionThreeList.Remove(currentStudent);
                        currentDistrictClassCount.Count--;
                    }
                }

            }

            stillOneBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionOneCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });
            stillTwoBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionTwoCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });
            stillThreeBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionThreeCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });

            if (outOfClass)
            {
                currentStudent.SessionOne = exceedsCount;
                currentStudent.SessionTwo = exceedsCount;
                currentStudent.SessionThree = exceedsCount;
            }
            else
            {
                currentStudent.SessionOne = "";
                currentStudent.SessionTwo = "";
                currentStudent.SessionThree = "";
            }
        }

        /// <summary>
        /// Export workshop information to to XLSX file.
        /// </summary>
        public void createWorkshopExportExcel()
        {
            int columnCount = 15;
            FileInfo currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Workshop Final Schedule.xlsx");
            if (currentFile.Exists)
            {
                currentFile.Delete();  // ensures we create a new workbook
                currentFile = new FileInfo(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\Workshop Final Schedule.xlsx");
            }

            using (ExcelPackage currentExcel = new ExcelPackage(currentFile))
            {
                ExcelWorksheet currentSheet = currentExcel.Workbook.Worksheets.Add("Student Registrations");
                District studentDistrict;
                School studentSchool;
                Presenter workshopPresenter;
                int currentRowCounter = 2;

                currentSheet.Cells["A1"].Value = "Student First Name";
                currentSheet.Cells["B1"].Value = "Student Last Name";
                currentSheet.Cells["C1"].Value = "School Name";
                currentSheet.Cells["D1"].Value = "District Name";
                currentSheet.Cells["E1"].Value = "Student EMail";
                currentSheet.Cells["F1"].Value = "Student Phone";
                currentSheet.Cells["G1"].Value = "Student Grade Level";
                currentSheet.Cells["H1"].Value = "Student Number";
                currentSheet.Cells["I1"].Value = "Teacher Name";
                currentSheet.Cells["J1"].Value = "Session One Title";
                currentSheet.Cells["K1"].Value = "Session One Room";
                currentSheet.Cells["L1"].Value = "Session Two Title";
                currentSheet.Cells["M1"].Value = "Session Two Room";
                currentSheet.Cells["N1"].Value = "Session Three Title";
                currentSheet.Cells["O1"].Value = "Session Three Room";

                currentSheet.Cells["A1"].AutoFitColumns();


                String headerRange = "A1:" + Convert.ToChar('A' + columnCount - 1) + 1;

                using (ExcelRange currentRange = currentSheet.Cells[headerRange])
                {
                    currentRange.Style.WrapText = false;
                    currentRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    currentRange.Style.Font.Bold = true;
                    currentRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    currentRange.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    currentRange.Style.Font.Color.SetColor(Color.White);

                }

                ExcelWorksheet districtCountSheet = currentExcel.Workbook.Worksheets.Add("District Counts");
                createDistrictSheet(districtCountSheet);





                ExcelWorksheet sessionCountSheet = currentExcel.Workbook.Worksheets.Add("Session Counts");
                sessionCountSheet.Cells[1, 1].Value = "Title";
                sessionCountSheet.Cells[1, 2].Value = "Session A Count";
                sessionCountSheet.Cells[1, 3].Value = "Session A Left";
                sessionCountSheet.Cells[1, 4].Value = "Session B Count";
                sessionCountSheet.Cells[1, 5].Value = "Session B Left";
                sessionCountSheet.Cells[1, 6].Value = "Session C Count";
                sessionCountSheet.Cells[1, 7].Value = "Session C Left";

                for (int rowCount = 1; rowCount <= sessionList.Count; rowCount++)
                {
                    sessionCountSheet.Cells[rowCount + 1, 1].Value = sessionList[rowCount - 1].SessionPresenter.PresenterTitle;
                    sessionCountSheet.Cells[rowCount + 1, 2].Value = sessionList[rowCount - 1].SessionOneCount;
                    sessionCountSheet.Cells[rowCount + 1, 3].Value = sessionList[rowCount - 1].Capacity - sessionList[rowCount - 1].SessionOneCount;
                    sessionCountSheet.Cells[rowCount + 1, 4].Value = sessionList[rowCount - 1].SessionTwoCount;
                    sessionCountSheet.Cells[rowCount + 1, 5].Value = sessionList[rowCount - 1].Capacity - sessionList[rowCount - 1].SessionTwoCount;
                    sessionCountSheet.Cells[rowCount + 1, 6].Value = sessionList[rowCount - 1].SessionThreeCount;
                    sessionCountSheet.Cells[rowCount + 1, 7].Value = sessionList[rowCount - 1].Capacity - sessionList[rowCount - 1].SessionThreeCount;
                }

                #region Create master schedule sheet
                foreach (Student workshopStudent in studentList)
                {
                    studentSchool = schoolList.Find(delegate(School curr) { return curr.SchoolName.Equals(workshopStudent.SchoolName); });
                    workshopPresenter = presenterList.Find(delegate(Presenter curr) { return curr.PresenterTitle.Equals(workshopStudent.SessionOne); });

                    studentDistrict = studentSchool.SchoolDistrict;

                    currentSheet.Cells[currentRowCounter, 1].Value = workshopStudent.FirstName;
                    currentSheet.Cells[currentRowCounter, 2].Value = workshopStudent.LastName;
                    currentSheet.Cells[currentRowCounter, 3].Value = workshopStudent.SchoolName;
                    currentSheet.Cells[currentRowCounter, 4].Value = studentDistrict.DistrictName;
                    currentSheet.Cells[currentRowCounter, 5].Value = workshopStudent.EmailAddress;
                    currentSheet.Cells[currentRowCounter, 6].Value = workshopStudent.PhoneNumber;
                    currentSheet.Cells[currentRowCounter, 7].Value = workshopStudent.GradeLevel;
                    currentSheet.Cells[currentRowCounter, 8].Value = workshopStudent.StudentNumber;
                    currentSheet.Cells[currentRowCounter, 9].Value = workshopStudent.StudentTeacher;

                    if (workshopStudent.SessionOne.Equals(failedCutoff) || workshopStudent.SessionOne.Equals(exceedsCount))
                    {
                        currentSheet.Cells[currentRowCounter, 10].Value = workshopStudent.SessionOne;
                        currentSheet.Cells[currentRowCounter, 11].Value = workshopStudent.SessionOne;
                    }
                    else
                    {
                        currentSheet.Cells[currentRowCounter, 10].Value = workshopStudent.SessionOne;
                        workshopPresenter = presenterList.Find(delegate(Presenter curr) { return curr.PresenterTitle.Equals(workshopStudent.SessionOne); });
                        currentSheet.Cells[currentRowCounter, 11].Value = workshopPresenter.PresenterRoom.RoomName;
                    }
                    if (workshopStudent.SessionTwo.Equals(failedCutoff) || workshopStudent.SessionTwo.Equals(exceedsCount))
                    {
                        currentSheet.Cells[currentRowCounter, 12].Value = workshopStudent.SessionTwo;
                        currentSheet.Cells[currentRowCounter, 13].Value = workshopStudent.SessionTwo;
                    }
                    else
                    {
                        currentSheet.Cells[currentRowCounter, 12].Value = workshopStudent.SessionTwo;
                        workshopPresenter = presenterList.Find(delegate(Presenter curr) { return curr.PresenterTitle.Equals(workshopStudent.SessionTwo); });
                        currentSheet.Cells[currentRowCounter, 13].Value = workshopPresenter.PresenterRoom.RoomName;
                    }
                    if (workshopStudent.SessionThree.Equals(failedCutoff) || workshopStudent.SessionThree.Equals(exceedsCount))
                    {
                        currentSheet.Cells[currentRowCounter, 14].Value = workshopStudent.SessionThree;
                        currentSheet.Cells[currentRowCounter, 15].Value = workshopStudent.SessionThree;
                    }
                    else
                    {
                        currentSheet.Cells[currentRowCounter, 14].Value = workshopStudent.SessionThree;
                        workshopPresenter = presenterList.Find(delegate(Presenter curr) { return curr.PresenterTitle.Equals(workshopStudent.SessionThree); });
                        currentSheet.Cells[currentRowCounter, 15].Value = workshopPresenter.PresenterRoom.RoomName;
                    }

                    currentRowCounter++;

                }
                #endregion

                String rowsCellRange = "A2:" + Convert.ToChar('A' + columnCount - 1) + (studentList.Count + 1);
                using (ExcelRange currentRange = currentSheet.Cells[rowsCellRange])
                {
                    currentRange.Style.WrapText = true;
                    currentRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                }

                String currentPresenterTitle;
                int currentRowCounterA = 2, currentRowCounterB = 2, currentRowCounterC = 2;
                int roomNumber = 1;
                #region Create presenter rolls with attending student names
                foreach (Presenter currentPresenter in presenterList)
                {
                    Session currentSession = sessionList.Find(delegate(Session curr) { return (curr.SessionPresenter.PresenterTitle.Equals(currentPresenter.PresenterTitle)); });

                    if (currentPresenter.PresenterTitle.Length > 28)
                    {
                        currentPresenterTitle = currentPresenter.PresenterTitle.Substring(0, 23) + roomNumber;
                        roomNumber++;
                    }
                    else
                    {
                        currentPresenterTitle = currentPresenter.PresenterTitle;
                    }

                    ExcelWorksheet currentSheetA = currentExcel.Workbook.Worksheets.Add(currentPresenterTitle + " A");
                    ExcelWorksheet currentSheetB = currentExcel.Workbook.Worksheets.Add(currentPresenterTitle + " B");
                    ExcelWorksheet currentSheetC = currentExcel.Workbook.Worksheets.Add(currentPresenterTitle + " C");

                    currentPresenterTitle = currentPresenter.PresenterTitle;

                    currentSheetA.Cells["A1"].Value = currentPresenterTitle;
                    currentSheetA.Cells["B1"].Value = "Session A";
                    currentSheetA.Cells["C1"].Value = "Session Count: " + currentSession.SessionOneCount;

                    currentSheetB.Cells["A1"].Value = currentPresenterTitle;
                    currentSheetB.Cells["B1"].Value = "Session B";
                    currentSheetB.Cells["C1"].Value = "Session Count: " + currentSession.SessionTwoCount;

                    currentSheetC.Cells["A1"].Value = currentPresenterTitle;
                    currentSheetC.Cells["B1"].Value = "Session C";
                    currentSheetC.Cells["C1"].Value = "Session Count: " + currentSession.SessionThreeCount;



                    foreach (Student currentStudent in studentList)
                    {

                        if (currentStudent.SessionOne.Equals(currentPresenterTitle))
                        {
                            //write student name to currentsheetA
                            currentSheetA.Cells[currentRowCounterA, 1].Value = currentStudent.FirstName;
                            currentSheetA.Cells[currentRowCounterA, 2].Value = currentStudent.LastName;
                            currentRowCounterA++;
                        }
                        if (currentStudent.SessionTwo.Equals(currentPresenterTitle))
                        {
                            //write student name to currentsheetB
                            currentSheetB.Cells[currentRowCounterB, 1].Value = currentStudent.FirstName;
                            currentSheetB.Cells[currentRowCounterB, 2].Value = currentStudent.LastName;
                            currentRowCounterB++;
                        }
                        if (currentStudent.SessionThree.Equals(currentPresenterTitle))
                        {
                            //write student name to currentsheetC
                            currentSheetC.Cells[currentRowCounterC, 1].Value = currentStudent.FirstName;
                            currentSheetC.Cells[currentRowCounterC, 2].Value = currentStudent.LastName;
                            currentRowCounterC++;
                        }


                    }

                    //Reset row counters
                    currentRowCounterA = 2;
                    currentRowCounterB = 2;
                    currentRowCounterC = 2;
                }
                #endregion
                currentExcel.Save();
            }
            okToSchedule = true;
        }

        //stillOneBlank = sessionList.FindAll(delegate(Session tempSession) { return (tempSession.SessionOneCount < tempSession.SessionPresenter.PresenterRoom.RoomCapacity); });
       /// <summary>
       /// Create the XLSX sheet with district information for export.
       /// </summary>
       /// <param name="districtCountSheet"></param>
        private void createDistrictSheet(ExcelWorksheet districtCountSheet)
        {
            districtCountSheet.Cells[1, 1].Value = "Session Title";
            districtCountSheet.Cells[1, 2].Value = "District";
            districtCountSheet.Cells[1, 3].Value = "Total district students in session";
            districtCountSheet.Cells[1, 4].Value = "District allowed amount for session";

            int districtStudentCount = 0;
            School currentSchool = null;
            District currentDistrict = null;


            for (int sessionCount = 0; sessionCount < sessionList.Count; sessionCount++)
            {
                for (int districtCount = 0; districtCount < districtList.Count; districtCount++)
                {
                    District testDistrict = districtList[districtCount];
                    for (int studentCount = 0; studentCount < sessionList[sessionCount].SessionOneList.Count; studentCount++)
                    {
                        Student currentStudent = sessionList[sessionCount].SessionOneList[studentCount];
                        currentSchool = schoolList.Find(delegate(School tempSchool) { return (tempSchool.SchoolName.Equals(currentStudent.SchoolName)); });
                        currentDistrict = currentSchool.SchoolDistrict;

                        if (currentDistrict.DistrictName.Equals(testDistrict.DistrictName))
                        {
                            districtStudentCount++;
                        }

                    }

                    for (int studentCount = 0; studentCount < sessionList[sessionCount].SessionTwoList.Count; studentCount++)
                    {
                        Student currentStudent = sessionList[sessionCount].SessionTwoList[studentCount];
                        currentSchool = schoolList.Find(delegate(School tempSchool) { return (tempSchool.SchoolName.Equals(currentStudent.SchoolName)); });
                        currentDistrict = currentSchool.SchoolDistrict;

                        if (currentDistrict.DistrictName.Equals(testDistrict.DistrictName))
                        {
                            districtStudentCount++;
                        }

                    }

                    for (int studentCount = 0; studentCount < sessionList[sessionCount].SessionThreeList.Count; studentCount++)
                    {
                        Student currentStudent = sessionList[sessionCount].SessionThreeList[studentCount];
                        currentSchool = schoolList.Find(delegate(School tempSchool) { return (tempSchool.SchoolName.Equals(currentStudent.SchoolName)); });
                        currentDistrict = currentSchool.SchoolDistrict;

                        if (currentDistrict.DistrictName.Equals(testDistrict.DistrictName))
                        {
                            districtStudentCount++;
                        }

                    }

                    int insertRow = (sessionCount * districtList.Count) + districtCount + 2;
                    districtCountSheet.Cells[insertRow, 1].Value = sessionList[sessionCount].SessionPresenter.PresenterTitle;
                    districtCountSheet.Cells[insertRow, 2].Value = testDistrict.DistrictName;
                    districtCountSheet.Cells[insertRow, 3].Value = districtStudentCount;
                    districtCountSheet.Cells[insertRow, 4].Value = (testDistrict.DistrictPercentage / 100.00) * (3 * sessionList[sessionCount].SessionPresenter.PresenterRoom.RoomCapacity);
                    districtStudentCount = 0;
                }






            }
        }

        /// <summary>
        /// Starts the scheduling program.  
        /// </summary>
        public void startSchedule()
        {
            importExcelDataToLists();
        }
    }
}
