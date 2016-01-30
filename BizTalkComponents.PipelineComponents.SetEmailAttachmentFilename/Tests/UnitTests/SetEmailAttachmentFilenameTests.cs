using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Winterdom.BizTalk.PipelineTesting;
using BizTalkComponents.Utils;

namespace BizTalkComponents.PipelineComponents.SetEmailAttachmentFilename.Tests.UnitTests
{
    [TestClass]
    public class SetEmailAttachmentFilenameTests
    {
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void IncorrectParametersTest()
        {
            var pipeline = PipelineFactory.CreateEmptySendPipeline();
            var useMessageFilename = false;
            var filename = "testmessage1";

            var component = new SetEmailAttachmentFilename
            {
                UseMessageFilename = useMessageFilename,
            };

            pipeline.AddComponent(component, PipelineStage.Encode);
            var message = MessageHelper.Create("<TestMessage1><Name>namnet</Name></TestMessage1>");
            ContextExtensions.Write(message.Context, new ContextProperty(FileProperties.ReceivedFileName), filename);
            var result = pipeline.Execute(message);
            Assert.IsTrue(true);
        }



        [TestMethod]
        public void UseMessageExtensionTest()
        {
            var pipeline = PipelineFactory.CreateEmptySendPipeline();
            var useMessageFilename = true;
            var filename = "testmessage1.xml";

            var component = new SetEmailAttachmentFilename
            {
                UseMessageFilename = useMessageFilename,
            };

            pipeline.AddComponent(component, PipelineStage.Encode);
            var message = MessageHelper.Create("<TestMessage1><Name>namnet</Name></TestMessage1>");
            ContextExtensions.Write(message.Context, new ContextProperty(FileProperties.ReceivedFileName), filename);
            var result = pipeline.Execute(message);
            Assert.IsTrue(true);
        }

        [TestMethod]
        public void FileExtensionTest()
        {
            var pipeline = PipelineFactory.CreateEmptySendPipeline();
            var useMessageFilename = false;
            var filename = "testmessage1";
            var fileExtension = ".xml";

            var component = new SetEmailAttachmentFilename
            {
                UseMessageFilename = useMessageFilename,
                FileExtension = fileExtension
            };

            pipeline.AddComponent(component, PipelineStage.Encode);
            var message = MessageHelper.Create("<TestMessage1><Name>namnet</Name></TestMessage1>");
            ContextExtensions.Write(message.Context, new ContextProperty(FileProperties.ReceivedFileName), filename);
            var result = pipeline.Execute(message);
            Assert.IsTrue(true);
        }
    }
}
