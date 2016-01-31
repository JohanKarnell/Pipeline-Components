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
        public void UseReceivedFilenameTest()
        {
            var pipeline = PipelineFactory.CreateEmptySendPipeline();
            var useMessageFilename = true;
            var filename = "testmessage1.xml";
            var component = new SetEmailAttachmentFilename
            {
                 PropertyPath = "http://schemas.microsoft.com/BizTalk/2003/file-properties#ReceivedFileName"
            };
            pipeline.AddComponent(component, PipelineStage.Encode);
            var message = MessageHelper.Create("<TestMessage1><Name>namnet</Name></TestMessage1>");
            ContextExtensions.Write(message.Context, new ContextProperty(FileProperties.ReceivedFileName), filename);
            var result = pipeline.Execute(message);
            Assert.IsTrue(true);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void UseReceivedFilenameArgumentExceptionTest()
        {
            var pipeline = PipelineFactory.CreateEmptySendPipeline();
            var useMessageFilename = true;
            var filename = "testmessage1.xml";
            var component = new SetEmailAttachmentFilename
            {
                PropertyPath = "INCORRECTFORMAT"
            };
            pipeline.AddComponent(component, PipelineStage.Encode);
            var message = MessageHelper.Create("<TestMessage1><Name>namnet</Name></TestMessage1>");
            ContextExtensions.Write(message.Context, new ContextProperty(FileProperties.ReceivedFileName), filename);
            var result = pipeline.Execute(message);
            Assert.IsTrue(true);
        }
    }
}
