using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WFC_Scheduler.Model
{
    class Room
    {
        private String roomName;
        private int roomCapacity;

        public String RoomName
        {
            get { return roomName; }
            set { roomName = value; }
        }

        public int RoomCapacity
        {
            get { return roomCapacity; }
            set { roomCapacity = value; }
        }

        public Room(string roomName, int roomCapacity)
        {
            this.roomCapacity = roomCapacity;
            this.roomName = roomName;
        }

    }
}
