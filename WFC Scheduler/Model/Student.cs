using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace WFC_Scheduler.Model
{
    class Student
    {
        private String firstName, lastName, studentTeacher, studentNumber, emailAddress, phoneNumber, gradeLevel, schoolName, sessionOne, sessionTwo, sessionThree;
        private Request studentRequest;
        private bool okToSchedule;

        public bool OkToSchedule
        {
            get { return okToSchedule; }
            set { okToSchedule = value; }
        }

        public String FirstName
        {
            get { return firstName; }
            set { firstName = value; }
        }

        public String LastName
        {
            get { return lastName; }
            set { lastName = value; }
        }
        
        public String StudentTeacher
        {
            get { return studentTeacher; }
            set { studentTeacher = value; }
        }

        public String StudentNumber
        {
            get { return studentNumber; }
            set { studentNumber = value; }
        }

        public String EmailAddress
        {
            get { return emailAddress; }
            set { emailAddress = value; }
        }

        public String PhoneNumber
        {
            get { return phoneNumber; }
            set { phoneNumber = value; }
        }

        public String SchoolName
        {
            get { return schoolName; }
            set { schoolName = value; }
        }

        public String GradeLevel
        {
            get { return gradeLevel; }
            set { gradeLevel = value; }
        }

        public String SessionOne
        {
            get { return sessionOne; }
            set { sessionOne = value; }
        }

        public String SessionTwo
        {
            get { return sessionTwo; }
            set { sessionTwo = value; }
        }

        public String SessionThree
        {
            get { return sessionThree; }
            set { sessionThree = value; }
        }

        public Request StudentRequest
        {
            get { return studentRequest; }
            set { studentRequest = value; }
        }

       
        
        public Student(String firstName, String lastName, String studentNumber, String emailAddress, String studentTeacher, String phoneNumber, String gradeLevel, string schoolName, bool okToSchedule)
        {
            this.firstName = firstName;
            this.lastName = lastName;
            this.studentTeacher = studentTeacher;
            this.studentNumber = studentNumber;
            this.emailAddress = emailAddress;
            this.phoneNumber = phoneNumber;
            this.gradeLevel = gradeLevel;
            this.schoolName = schoolName;
            sessionOne = "";
            sessionTwo = "";
            sessionThree = "";
            this.okToSchedule = okToSchedule;
        }
    }
}
