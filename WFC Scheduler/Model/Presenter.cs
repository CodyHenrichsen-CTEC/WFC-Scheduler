﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WFC_Scheduler.Model
{
    class Presenter
    {
        private String presenterTitle, presenterDescription;
        private Room presenterRoom;

        public String PresenterTitle
        {
            get { return presenterTitle; }
            set { presenterTitle = value; }
        }

        public String PresenterDescription
        {
            get { return presenterDescription; }
            set { presenterDescription = value; }
        }

        public Room PresenterRoom
        {
            get { return presenterRoom; }
            set { presenterRoom = value; }
        }

        public Presenter(string presenterTitle, string presenterDescription, Room presenterRoom)
        {
            this.presenterDescription = presenterDescription;
            this.presenterTitle = presenterTitle;
            this.presenterRoom = presenterRoom;
        }

    }
}
