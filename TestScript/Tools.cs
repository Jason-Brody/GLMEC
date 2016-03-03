using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.ComponentModel.DataAnnotations;
using System.IO;

namespace TestScript
{
    public class Tools
    {
        public static void MasterDataVerification<T>(T data) where T : class
        {
            if (data == null)
                throw new ArgumentNullException(typeof(T).Name);

            var props = typeof(T).GetProperties();

            foreach (var prop in props)
            {
                var require = prop.GetCustomAttribute<RequiredAttribute>();
                if (require != null)
                {
                    object value = prop.GetValue(data);
                    var str = value?.ToString();
                    if (string.IsNullOrEmpty(str))
                    {
                        var alias = prop.GetCustomAttribute<DisplayAttribute>();
                        throw new ArgumentNullException(alias == null ? prop.Name : alias.Name);
                    }
                }
            }

        }

        public static List<string> GetDatas(string filePath)
        {
            List<string> datas = new List<string>();
            using (StreamReader sr = new StreamReader(filePath))
            {
                while (!sr.EndOfStream)
                {
                    datas.Add( sr.ReadLine());
                }

            }
            return datas;
        }
    }
}
