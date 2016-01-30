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

namespace BizTalkComponents.PipelineComponents.SetEmailAttachmentFilename
{
    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [ComponentCategory(CategoryTypes.CATID_Encoder)]
    [System.Runtime.InteropServices.Guid("9d0e4103-4cce-4536-83fa-4a5040674ad6")]
    public partial class SetEmailAttachmentFilename : IBaseComponent, IComponent, IComponentUI, IPersistPropertyBag
    {
        private const string FileExtensionPropertyName = "FileExtension";
        private const string UseMessageFilenamePropertyName = "UseMessageFilename";
        private const string FilenamePropertyName = "Filename";
        private const string XPathsPropertyName = "XPaths";
        private const string SeparatorCharPropertyName = "SeparatorChar";

        #region Parameters
        
        [DisplayName("File Extension")]
        [Description("The file extension to set on the email attachment")]
        public string FileExtension { get; set; }

        [DisplayName("Filename")]
        [Description("The filename to set for the attached message. Do not include any file extension here")]
        public string Filename { get; set; }

        [DisplayName("Xpaths")]
        [Description("Xpaths to message records. Separate multiple xpaths with the '|' char")]
        public string XPaths { get; set; }

        [DisplayName("SeparatorChar")]
        [Description("Charachter to use as separater in the filename between the Xpath record values")]
        public string SeparatorChar { get; set; }

        [RequiredRuntime]
        [DisplayName("Use message filename")]
        [Description("Use the processed message filename")]
        public bool UseMessageFilename { get; set; }

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
                if (UseMessageFilename)
                {
                    var receivedFileName = ContextExtensions.Read(pInMsg.Context, new ContextProperty(FileProperties.ReceivedFileName));
                    pInMsg.BodyPart.PartProperties.Write("FileName", "http://schemas.microsoft.com/BizTalk/2003/mime-properties", receivedFileName);
                }

                else if (!String.IsNullOrEmpty(FileExtension))
                {
                    pInMsg.BodyPart.PartProperties.Write("FileName", "http://schemas.microsoft.com/BizTalk/2003/mime-properties", FileExtension);
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
            UseMessageFilename = PropertyBagHelper.ReadPropertyBag<bool>(propertyBag, UseMessageFilenamePropertyName);
            Filename = PropertyBagHelper.ReadPropertyBag<string>(propertyBag, FilenamePropertyName);
            XPaths = PropertyBagHelper.ReadPropertyBag<string>(propertyBag, XPathsPropertyName);
            SeparatorChar = PropertyBagHelper.ReadPropertyBag<string>(propertyBag, SeparatorCharPropertyName);
        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
            PropertyBagHelper.WritePropertyBag(propertyBag, FileExtensionPropertyName, FileExtension);
            PropertyBagHelper.WritePropertyBag(propertyBag, UseMessageFilenamePropertyName, UseMessageFilename);
            PropertyBagHelper.WritePropertyBag(propertyBag, FilenamePropertyName, Filename);
            PropertyBagHelper.WritePropertyBag(propertyBag, XPathsPropertyName, XPaths);
            PropertyBagHelper.WritePropertyBag(propertyBag, SeparatorCharPropertyName, SeparatorChar);
        }
    }
}
