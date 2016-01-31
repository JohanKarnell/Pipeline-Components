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
        public string Name { get { return "SetEmailAttachmentFilename"; } }
        public string Version { get { return "1.0"; } }
        public string Description { get { return "Sets the filename of an email attachment sent by BizTalk via the SMTP adapter by picking the value from a promoted promoted property"; } }

        public IntPtr Icon
        {
            get { return IntPtr.Zero; }
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
            classID = new Guid("7A638E15-52B8-4F77-80B4-A248887FF44B");
        }

        public void InitNew()
        {
        }

    }
}
