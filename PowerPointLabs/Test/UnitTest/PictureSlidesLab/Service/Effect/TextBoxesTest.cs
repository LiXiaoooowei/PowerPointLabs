﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using Test.Util;

namespace Test.UnitTest.PictureSlidesLab.Service.Effect
{
    [TestClass]
    public class TextBoxesTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "PictureSlidesLab\\TextBoxes.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxInfoForEmptyTextBoxes()
        {
            var shapes = PpOperations.SelectShapesByPrefix("TextBox");
            var textBoxes = new TextBoxes(shapes, 
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight);
            var textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.AreEqual(null, textBoxInfo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxInfo()
        {
            PpOperations.SelectSlide(2);
            var shapes = PpOperations.SelectShapesByPrefix("TextBox");
            var textBoxes = new TextBoxes(shapes,
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight);
            var textBoxInfo = textBoxes.GetTextBoxesInfo();

            Assert.IsTrue(SlideUtil.IsRoughlySame(348.4239f, textBoxInfo.Height));
            Assert.IsTrue(SlideUtil.IsRoughlySame(68.2f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(52.17752f, textBoxInfo.Top));
            Assert.IsTrue(SlideUtil.IsRoughlySame(741.0565f, textBoxInfo.Width));

            TextBoxes.AddMargin(textBoxInfo, 25);
            Assert.IsTrue(SlideUtil.IsRoughlySame(398.4239f, textBoxInfo.Height));
            Assert.IsTrue(SlideUtil.IsRoughlySame(43.1999969f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(27.17752f, textBoxInfo.Top));
            Assert.IsTrue(SlideUtil.IsRoughlySame(791.0565f, textBoxInfo.Width));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStartBoxing()
        {
            PpOperations.SelectSlide(2);
            var shapes = PpOperations.SelectShapesByPrefix("TextBox");
            var textBoxes = new TextBoxes(shapes,
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight);

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Left)
                .StartBoxing();
            var textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(241.439987f, textBoxInfo.Height));
            Assert.IsTrue(SlideUtil.IsRoughlySame(25f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(149.279953f, textBoxInfo.Top));
            Assert.IsTrue(SlideUtil.IsRoughlySame(710.945068f, textBoxInfo.Width));

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Centre)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(124.527481f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(149.279953f, textBoxInfo.Top));

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.BottomLeft)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(25f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(273.559875f, textBoxInfo.Top));

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Bottom)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(124.527481f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(273.560028f, textBoxInfo.Top));

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Right)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(224.054962f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(149.279953f, textBoxInfo.Top));

            textBoxes
                .SetAlignment(Alignment.Auto)
                .SetPosition(Position.Original)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(68.2f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(52.17752f, textBoxInfo.Top));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStartBoxingWithTextWrapping()
        {
            PpOperations.SelectSlide(2);
            var shapes = PpOperations.SelectShapesByPrefix("TextBox");
            var textBoxes = new TextBoxes(shapes,
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight);

            textBoxes.StartTextWrapping();

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Left)
                .StartBoxing();

            var textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(349.440033f, textBoxInfo.Height));
            Assert.IsTrue(SlideUtil.IsRoughlySame(25f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(95.27996f, textBoxInfo.Top));
            // aft text wrapping, width is smaller (originally should be 710)
            Assert.IsTrue(SlideUtil.IsRoughlySame(448.355042f, textBoxInfo.Width));
        }
    }
}