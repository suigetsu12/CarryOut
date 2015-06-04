using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksInteropLibrary
{
    public enum SyncStatus
    {
        InProgress,
        Completed
    }
    public static class SyncManager
    {
        static Dictionary<string, string> _data = new Dictionary<string, string>();

        public static void StartSync(string listId)
        {
            if(_data.ContainsKey(listId))
            {
                _data[listId] = SyncStatus.InProgress.ToString();
            }
            else
            {
                _data.Add(listId, SyncStatus.InProgress.ToString());
            }
        }

        public static void CompleteSync(string listId)
        {
            if (_data.ContainsKey(listId))
            {
                _data[listId] = SyncStatus.Completed.ToString();
            }
            else
            {
                _data.Add(listId, SyncStatus.Completed.ToString());
            }
        }

        public static bool IsInprogressSync(string listId)
        {
            var result = false;
            if(_data.ContainsKey(listId))
            {
                var enumStatus = SyncStatus.Completed;
                var status = Enum.TryParse(_data[listId].ToString(), out enumStatus);
                if (enumStatus == SyncStatus.Completed)
                {
                    result = false;
                }
                else
                {
                    result = true;
                }
            }

            return result;
        }
    }
}
