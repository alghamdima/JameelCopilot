using System;
using System.IO;
using JameelCopilot.Services;
using JameelCopilot.Tests.Helpers;
using Xunit;

namespace JameelCopilot.Tests.Services
{
    /// <summary>The Jameel Copilot mark as the add-in carries it: one 256 px picture, scaled on demand, and an icon file.</summary>
    public class LogoImagesTests
    {
        [Fact]
        public void Master_Is256Square_WithTransparentCorners()
        {
            using (var master = LogoImages.Master())
            {
                Assert.Equal((256, 256), (master.Width, master.Height));
                Assert.Equal(0, master.GetPixel(1, 1).A);
                Assert.Equal(0, master.GetPixel(254, 254).A);
            }
        }

        [Fact]
        public void Master_IsTheTealEmblem_OnNoBackground()
        {
            using (var master = LogoImages.Master())
            {
                // The pentagon's top edge.
                var edge = master.GetPixel(128, 10);
                Assert.True(edge.A == 255 && edge.B > edge.G && edge.G > edge.R, edge.ToString());

                // Just inside that edge, and below the emblem: clear, not a filled tile.
                Assert.Equal(0, master.GetPixel(128, 30).A);
                Assert.Equal(0, master.GetPixel(128, 240).A);
            }
        }

        [Fact]
        public void Master_IsANewCopyEachTime()
        {
            var first = LogoImages.Master();
            first.Dispose();

            using (var second = LogoImages.Master())
            {
                Assert.Equal(256, second.Width);
            }
        }

        [Theory]
        [InlineData(16)]
        [InlineData(32)]
        [InlineData(48)]
        [InlineData(80)]
        public void Sized_HasTheSizeAsked(int pixels)
        {
            using (var image = LogoImages.Sized(pixels))
            {
                Assert.Equal((pixels, pixels), (image.Width, image.Height));
            }
        }

        [Theory]
        [InlineData(4, 16)]
        [InlineData(-1, 16)]
        [InlineData(5000, 256)]
        public void Sized_IsClamped(int asked, int expected)
        {
            using (var image = LogoImages.Sized(asked))
            {
                Assert.Equal(expected, image.Width);
            }
        }

        // Scaling must not fill the clear corners in: on the ribbon they would show as a box.
        [Fact]
        public void Sized_KeepsTheCornersClear_AndTheEmblemVisible()
        {
            using (var image = LogoImages.Sized(32))
            {
                Assert.True(image.GetPixel(0, 0).A < 64, "corner alpha " + image.GetPixel(0, 0).A);

                // The emblem's lines are thin, so at 32 px no pixel need be fully solid.
                var strongest = 0;
                for (var y = 0; y < 32; y++)
                    for (var x = 0; x < 32; x++)
                        strongest = Math.Max(strongest, image.GetPixel(x, y).A);
                Assert.True(strongest >= 192, "strongest alpha " + strongest);
            }
        }

        [Fact]
        public void Ico_HoldsEightImages_SmallestFirst()
        {
            var ico = File.ReadAllBytes(RepoFiles.Find("VSTO2", "JameelCopilot", "Resources", "jc.ico"));

            Assert.Equal(new byte[] { 0, 0, 1, 0, 8, 0 }, new[] { ico[0], ico[1], ico[2], ico[3], ico[4], ico[5] });
            Assert.Equal(16, ico[6]);
            // The last entry is the 256 px image, whose size is written as 0.
            Assert.Equal(0, ico[6 + 16 * 7]);
            // Each image is stored as a PNG.
            var firstOffset = System.BitConverter.ToInt32(ico, 6 + 12);
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, new[] { ico[firstOffset], ico[firstOffset + 1], ico[firstOffset + 2], ico[firstOffset + 3] });
        }
    }
}
