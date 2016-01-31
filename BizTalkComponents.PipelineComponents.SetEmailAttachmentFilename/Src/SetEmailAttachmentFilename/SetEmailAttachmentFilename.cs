using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BizTalkComponents.Utils;
using Microsoft.BizTalk.Component.Interop;
using Microsoft.BizTalk.Message.Interop;
using IComponent = Microsoft.BizTalk.Component.Interop.IComponent;
using System.ComponentModel;
using System.Diagnostics;
using System.ComponentModel.DataAnnotations;

namespace BizTalkComponents.PipelineComponents.SetEmailAttachmentFilename
{
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [ComponentCategory(CategoryTypes.CATID_Encoder)]
    [System.Runtime.InteropServices.Guid("9d0e4103-4cce-4536-83fa-4a5040674ad6")]
    public partial class SetEmailAttachmentFilename : IBaseComponent, IComponent, IComponentUI, IPersistPropertyBag
    {
        private const string PropertyPathPropertyName = "PropertyPath";

        [DisplayName("Property Path")]
        [Description("The property path where the specified value will be promoted to, i.e. http://temupuri.org#MyProperty.")]
        [RegularExpression(@"^.*#.*$", ErrorMessage = "A property path should be formatted as namespace#property.")]
        [RequiredRuntime]
        public string PropertyPath { get; set; }

        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            string errorMessage;

            if (!Validate(out errorMessage))
            {
                throw new ArgumentException(errorMessage);
            }

            try
            {
                var attachmentFileName = ContextExtensions.Read(pInMsg.Context, new ContextProperty(PropertyPath));
                pInMsg.BodyPart.PartProperties.Write("FileName", "http://schemas.microsoft.com/BizTalk/2003/mime-properties", attachmentFileName);
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("SetEmailAttachmentFilename", ex.Message, EventLogEntryType.Error);
                throw new Exception(ex.Message);
            }

            return pInMsg;
        }

        public void Load(IPropertyBag propertyBag, int errorLog)
        {
            PropertyPath = PropertyBagHelper.ReadPropertyBag(propertyBag, PropertyPathPropertyName, PropertyPath);
        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
            PropertyBagHelper.WritePropertyBag(propertyBag, PropertyPathPropertyName, PropertyPath);
        }
    }
}
