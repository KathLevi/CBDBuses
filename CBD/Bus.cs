using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CBD {
    class Bus {
        private int bus_number;
        private int bus_curr_on_board = 0;
        private int bus_capacity;
        List<Group> groups_on_bus;

        Bus() { }
        Bus(int _num, int _cap) {
            bus_number = _num;
            bus_capacity = _cap;
        }
        public int GetBusNum() { return bus_number; }
        public int GetBusCap() { return bus_capacity; }
        public int GetBusCurr() { return bus_curr_on_board; }
        public void AddToBusCurr(int add) { bus_curr_on_board += add; }
        public void SubToBusCurr(int sub) { bus_curr_on_board -= sub; }
        public bool AddGroupToBus(Group g) {
            if ((g.GetSizeOf() + GetBusCurr()) <= bus_capacity) {
                groups_on_bus.Add(g);
                AddToBusCurr(g.GetSizeOf());
                return true;
            }
            return false;
        }
        public bool RemoveGroupFromBus(Group g) {
            foreach(Group _g in groups_on_bus) {
                if (_g == g) {
                    groups_on_bus.Remove(g);
                    SubToBusCurr(g.GetSizeOf());
                    return true;
                }
            }
            return false;
        }
    }
}
