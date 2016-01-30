using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BizTalkComponents.Utils;

namespace BizTalkComponents.PipelineComponents.SetEmailAttachmentFilename
{
    public partial class SetEmailAttachmentFilename
    {
        public string Name { get { return "SetFileExtOnEmailAttachment"; } }
        public string Version { get { return "1.0"; } }
        public string Description { get { return "Sets the file extension of email attachments sent by BizTalk via the SMTP adapter"; } }

        public IntPtr Icon
        {
            get { throw new NotImplementedException(); }
        }

        public IEnumerator Validate(object projectSystem)
        {
            return ValidationHelper.Validate(this, false).ToArray().GetEnumerator();
        }

        public bool Validate(out string errorMessage)
        {
            var errors = ValidationHelper.Validate(this, true).ToArray();

            if (errors.Any())
            {
                errorMessage = string.Join(",", errors);

                return false;
            }

            errorMessage = string.Empty;

            return true;
        }

        public void GetClassID(out Guid classID)
        {
            classID = new Guid("0E3ACE0E-49FE-499C-8945-581A0CF8D8D3");
        }

        public void InitNew()
        {
        }

    }
}
