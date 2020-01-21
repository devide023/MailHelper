using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailHelper
{
    public class SendServerConfigurationEntity
    {
        public string ImapHost { get; set; }
        public int ImapPort { get; set; }
        /// <summary>
        /// 邮箱SMTP服务器地址
        /// </summary>
        public string SmtpHost { get; set; }

        /// <summary>
        /// 邮箱SMTP服务器端口
        /// </summary>
        public int SmtpPort { get; set; }

        /// <summary>
        /// 是否启用IsSsl
        /// </summary>
        public bool IsSsl { get; set; }

        /// <summary>
        /// 邮件编码
        /// </summary>
        public string MailEncoding { get; set; }

        /// <summary>
        /// 邮箱账号
        /// </summary>
        public string SenderAccount { get; set; }

        /// <summary>
        /// 邮箱密码
        /// </summary>
        public string SenderPassword { get; set; }
    }
}
