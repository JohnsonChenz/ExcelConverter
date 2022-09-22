using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using UnityEngine;
using UnityEngine.UIElements;

namespace ExcelConverter
{
    public static class LogAdder
    {
        private static ScrollView scrollViewLog;
        private static Queue<Label> logCache = new Queue<Label>();

        public static async void OnUpdate()
        {
            while (logCache.Count > 0)
            {
                Label label = logCache.Dequeue();
                scrollViewLog.Add(label);
                await Task.Yield();
                scrollViewLog.verticalScroller.ScrollPageDown(scrollViewLog.verticalScroller.highValue);
            }
        }

        public static void SetScrollView(ScrollView scrollViewLog)
        {
            LogAdder.scrollViewLog = scrollViewLog;
            ClearLog();
        }

        public static void AddLog(string text, string colorHex = "#FFED94")
        {
            if (scrollViewLog == null)
            {
                Debug.LogError("請設置Logger的ScrollView!!");
                return;
            }

            Label label = new Label("\n" + $"<color={colorHex}>" + text + "</color>");
            logCache.Enqueue(label);
        }

        public static void AddLogError(string text, string colorHex = "#FF0000")
        {
            if (scrollViewLog == null)
            {
                Debug.LogError("請設置Logger的ScrollView!!");
                return;
            }

            Label label = new Label("\n" + $"<color={colorHex}>" + text + "</color>");
            logCache.Enqueue(label);
        }

        public static Label AddLogWithGet()
        {
            if (scrollViewLog == null)
            {
                Debug.LogError("請設置Logger的ScrollView!!");
                return null;
            }

            Label label = new Label();
            logCache.Enqueue(label);
            return label;
        }

        public static void ClearLog()
        {
            scrollViewLog.Clear();
        }
    }
}
