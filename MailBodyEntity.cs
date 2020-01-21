using MimeKit.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    /// <summary>
    /// 邮件内容实体
    /// </summary>
    public class MailBodyEntity
    {
        ///// <summary>
        ///// 邮件文本内容
        ///// </summary>
        //public string MailTextBody { get; set; }

        /// <summary>
        /// 邮件内容类型
        /// </summary>
        public TextFormat MailBodyType { get; set; } = TextFormat.Html;

        /// <summary>
        /// 邮件附件集合
        /// </summary>
        public List<MailFile> MailFiles { get; set; }

        /// <summary>
        /// 收件人
        /// </summary>
        public List<string> Recipients { get; set; }

        /// <summary>
        /// 抄送
        /// </summary>
        public List<string> Cc { get; set; }

        /// <summary>
        /// 密送
        /// </summary>
        public List<string> Bcc { get; set; }

        /// <summary>
        /// 发件人
        /// </summary>
        public string Sender { get; set; }

        /// <summary>
        /// 发件人地址
        /// </summary>
        public string SenderAddress { get; set; }

        /// <summary>
        /// 邮件主题
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// 邮件内容
        /// </summary>
        public string Body { get; set; }
    }
}
