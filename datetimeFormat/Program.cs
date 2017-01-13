using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace datetimeFormat
{
    class Program
    {
        #region //DatetimeFormatPrint
        public static void DatetimeFormatPrint()
        {
            Console.WriteLine(string.Format("{0:R}", new DateTime(1, 1, 1)));

            //2008年4月24日
            Console.WriteLine(string.Format("D:{0}", DateTime.Now.ToString("D")));
            //2008-4-24
            Console.WriteLine(string.Format("d:{0}", DateTime.Now.ToString("d")));
            //2008年4月24日 16:30:15
            Console.WriteLine(string.Format("F:{0}", DateTime.Now.ToString("F")));
            //2008年4月24日 16:30
            Console.WriteLine(string.Format("f:{0}", DateTime.Now.ToString("f")));
            //2008-4-24 16:30:15
            Console.WriteLine(string.Format("G:{0}", DateTime.Now.ToString("G")));
            //2008-4-24 16:30
            Console.WriteLine(string.Format("g:{0}", DateTime.Now.ToString("g")));
            //16:30:15
            Console.WriteLine(string.Format("T:{0}", DateTime.Now.ToString("T")));
            //16:30
            Console.WriteLine(string.Format("t:{0}", DateTime.Now.ToString("t")));
            //2008年4月24日 8:30:15
            Console.WriteLine(string.Format("U:{0}", DateTime.Now.ToString("U")));
            //2008-04-24 16:30:15Z
            Console.WriteLine(string.Format("u:{0}", DateTime.Now.ToString("u")));
            //4月24日
            Console.WriteLine(string.Format("m:{0}", DateTime.Now.ToString("m")));
            Console.WriteLine(string.Format("M:{0}", DateTime.Now.ToString("M")));
            //Tue, 24 Apr 2008 16:30:15 GMT
            Console.WriteLine(string.Format("r:{0}", DateTime.Now.ToString("r")));
            Console.WriteLine(string.Format("R:{0}", DateTime.Now.ToString("R")));
            //2008年4月
            Console.WriteLine(string.Format("y:{0}", DateTime.Now.ToString("y")));
            Console.WriteLine(string.Format("Y:{0}", DateTime.Now.ToString("Y")));
            //2008-04-24T15:52:19.1562500+08:00
            Console.WriteLine(string.Format("o:{0}", DateTime.Now.ToString("o")));
            Console.WriteLine(string.Format("O:{0}", DateTime.Now.ToString("O")));
            //2008-04-24T16:30:15
            Console.WriteLine(string.Format("s:{0}", DateTime.Now.ToString("s")));
            //2008-04-24 15:52:19
            Console.WriteLine(string.Format("yyyy-MM-dd HH：mm：ss：ffff:{0}", DateTime.Now.ToString("yyyy-MM-dd HH：mm：ss：ffff")));
            //2008年04月24 15时56分48秒
            Console.WriteLine(string.Format("yyyy年MM月dd HH时mm分ss秒:{0}", DateTime.Now.ToString("yyyy年MM月dd HH时mm分ss秒")));
            //星期二, 四月 24 2008
            Console.WriteLine(string.Format("dddd, MMMM dd yyyy:{0}", DateTime.Now.ToString("dddd, MMMM dd yyyy")));
            //二, 四月 24 ’08
            Console.WriteLine(string.Format("ddd, MMM d \"’\"yy:{0}", DateTime.Now.ToString("ddd, MMM d \"’\"yy")));
            //星期二, 四月 24
            Console.WriteLine(string.Format("dddd, MMMM dd:{0}", DateTime.Now.ToString("dddd, MMMM dd")));
            //4-08
            Console.WriteLine(string.Format("M/yy:{0}", DateTime.Now.ToString("M/yy")));
            //24-04-08
            Console.WriteLine(string.Format("dd-MM-yy:{0}", DateTime.Now.ToString("dd-MM-yy")));
        }
        #endregion
        #region //StringFormatPrint
        public static void StringFormatPrint()
        {
            Console.WriteLine(string.Format("n:{0}", 12345.ToString("n")));  //生成 12,345.00
            Console.WriteLine(string.Format("C:{0}", 12345.ToString("C"))); //生成 ￥12,345.00
            Console.WriteLine(string.Format("e:{0}", 12345.ToString("e"))); //生成 1.234500e+004
            Console.WriteLine(string.Format("f4:{0}", 12345.ToString("f4"))); //生成 12345.0000
            Console.WriteLine(string.Format("x(16 Base):{0}", 12345.ToString("x"))); //生成 3039 (16进制)
            Console.WriteLine(string.Format("p:{0}", 12345.ToString("p"))); //生成 1,234,500%
        }
        #endregion
        #region //DatetimeCommonTransform
        public static void DatetimeCommonTransform()
        {
            //Today
            Console.WriteLine(string.Format("DateTime.Now.Date.ToShortDateString:{0}", DateTime.Now.Date.ToShortDateString()));
            Console.WriteLine(string.Format("DateTime.Now.DayOfWeek:{0}", DateTime.Now.DayOfWeek.ToString()));
            //Console.WriteLine(string.Format("ToInt16:{0}", DateTime.Now.AddDays(Convert.ToDouble((0 - Convert.ToInt16(DateTime.Now.DayOfWeek)))).ToShortDateString()));
            //Console.WriteLine(string.Format("ToInt16:{0}", DateTime.Now.AddDays(Convert.ToDouble((6 - Convert.ToInt16(DateTime.Now.DayOfWeek)))).ToShortDateString()));
            Console.WriteLine(string.Format("The last of the month:{0}", DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(1).AddDays(-1).ToShortDateString()));
            string[] Day = new string[] { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };
            Console.WriteLine(string.Format("The Chinese weekday:{0}", Day[Convert.ToInt16(DateTime.Now.DayOfWeek)]));
        }
        #endregion
        static void Main(string[] args)
        {
            //DatetimeFormatPrint();
            //StringFormatPrint();
            //DatetimeCommonTransform();
        }
    }
}