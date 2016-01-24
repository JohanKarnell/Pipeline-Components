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

namespace BizTalkComponents.PipelineComponents.SetFileExtOnEmailAttachment
{
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [ComponentCategory(CategoryTypes.CATID_Encoder)]
    [System.Runtime.InteropServices.Guid("9d0e4103-4cce-4536-83fa-4a5040674ad6")]
    public partial class SetFileExtOnEmailAttachment : IBaseComponent, IComponent, IComponentUI, IPersistPropertyBag
    {
        private const string FileExtensionPropertyName = "FileExtension";
        private const string UseMessageExtensionPropertyName = "UseMessageExtension";

        #region Parameters
        
        // Lägg till ytterligare en parameter? Output filename

        [DisplayName("File Extension")]
        [Description("The file extension to set on the email attachment")]
        public string FileExtension { get; set; }

        [RequiredRuntime]
        [DisplayName("Use message extension")]
        [Description("Use the processed message file extension")]
        public bool UseMessageExtension { get; set; }

        #endregion

        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            string errorMessage;

            if (!Validate(out errorMessage))
            {
                throw new ArgumentException(errorMessage);
            }

            try
            {
                // Filändelser sätts inte på utgående mail. Se över logik ytterligare
                // Enligt den URI ser jag ut att göra rätt: http://www.hivmr.com/db/ps7cpmx1dscxcxf8fzsp9kskfjzkfzs3
                if (UseMessageExtension)
                {
                    var recievedFileName = ContextExtensions.Read(pInMsg.Context, new ContextProperty(FileProperties.ReceivedFileName));
                    ContextExtensions.Write(pInMsg.Context, new ContextProperty(MIMEProperties.FileName), recievedFileName);
                }
                // Filändelser sätts inte på utgående mail. Se över logik ytterligare
                // Enligt den URI ser jag ut att göra rätt: http://www.hivmr.com/db/ps7cpmx1dscxcxf8fzsp9kskfjzkfzs3
                else if (!String.IsNullOrEmpty(FileExtension))
                {
                    ContextExtensions.Write(pInMsg.Context, new ContextProperty(MIMEProperties.FileName), FileExtension);
                }
                else
                {
                    string eventlogmessage = "The parameters was not set correctly on the pipeline. Please correct this and try again...";
                    EventLog.WriteEntry("SetFileExtOnEmailAttachment", eventlogmessage, EventLogEntryType.Error);
                    throw new ArgumentException(eventlogmessage);
                }
            }
            catch (ArgumentException ex)
            {
                EventLog.WriteEntry("SetFileExtOnEmailAttachment", ex.Message, EventLogEntryType.Error);
                throw new ArgumentException(ex.Message);
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("SetFileExtOnEmailAttachment", ex.Message, EventLogEntryType.Error);
                throw new Exception(ex.Message);
            }

            return pInMsg;
        }

        public void Load(IPropertyBag propertyBag, int errorLog)
        {
            FileExtension = PropertyBagHelper.ReadPropertyBag<string>(propertyBag, FileExtensionPropertyName);
            UseMessageExtension = PropertyBagHelper.ReadPropertyBag<bool>(propertyBag, UseMessageExtensionPropertyName);
        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
            PropertyBagHelper.WritePropertyBag(propertyBag, FileExtensionPropertyName, FileExtension);
            PropertyBagHelper.WritePropertyBag(propertyBag, UseMessageExtensionPropertyName, UseMessageExtension);
        }
    }
}
