using System.Collections.Generic;
using System.Runtime.InteropServices.Marshalling;

namespace AchieveItSystemModels
{
    public class ADailyWord
    {
        private string[] _aDailyWordLists;
        public string[] ADailyWordLists
        {
            get { return _aDailyWordLists; }
            set { _aDailyWordLists = value; }
        }
        /// <summary>
        /// 传入一串字符串，若是多个字符串在一起，请用 @ 分割！
        /// 字符串若单个 请不要含有 @ 字符
        /// </summary>
        /// <param name="stringLists"></param>
        public ADailyWord(string stringLists)
        {
            if (stringLists.Contains('@'))
            {
                ADailyWordLists = stringLists.Split('@');
            }
            else
            {
                int i = ADailyWordLists.Length - 1;
                ADailyWordLists[i] = stringLists;
            }
        }
    }
}