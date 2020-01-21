using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class MailFile
    {
        /// <summary>
        /// 邮件附件文件类型 例如：图片 MailFileType="image"
        /// </summary>
        public string MailFileType { get; set; }

        /// <summary>
        /// 邮件附件文件子类型 例如：图片 MailFileSubType="png"
        /// </summary>
        public string MailFileSubType { get; set; }

        /// <summary>
        /// 邮件附件文件路径  例如：图片 MailFilePath=@"C:\Files\123.png"
        /// </summary>
        public string MailFilePath { get; set; }
    }
}
