using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBD {
    class Group {
        private string group_name;
        private int group_size;

        Group() { }
        Group(string _name, int _size) {
            group_name = _name;
            group_size = _size;
        }
        public string GetGroupName() { return group_name; }
        public int GetSizeOf() { return group_size; }
        public static bool operator ==(Group g1, Group g2) {
            if (g1.GetGroupName() == g2.GetGroupName())
                return true;
            return false;
        }
        public static bool operator !=(Group g1, Group g2) {
            if (g1.GetGroupName() != g2.GetGroupName())
                return true;
            return false;
        }
    }
}
