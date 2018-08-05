namespace Core.Screenshot
{
    using System;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Windows.Forms;

    /// <summary>
    /// Provides functions to capture the entire screen, or a particular window, and save it to a file.
    /// </summary>
    public class ScreenCapture
    {
        private readonly ImageFormat _imageFormat;
        private readonly Rectangle _bounds;

        private const int SIZE_WIDTH = 1024;
        private const int SIZE_HEIGHT = 768;

        /// <summary>
        /// Constructor for ScreenCapture
        /// </summary>
        /// <param name="imageFormat">Image Format</param>
        public ScreenCapture(ImageFormat imageFormat)
        {
            _imageFormat = imageFormat ?? throw new Exception("Error, ImageFormat cannot be null.");
            _bounds = Screen.GetBounds(Point.Empty);
        }

        /// <summary>
        /// Constructor without parameters
        /// </summary>
        public ScreenCapture()
        {
            _imageFormat = ImageFormat.Jpeg;
            _bounds = Screen.GetBounds(Point.Empty);
        }

        /// <summary>
        /// Takes a fullscreen screenshot of the monitor and saves the specified file in a directory with custom name.
        /// It expects the Format of the file.
        /// </summary>
        /// </summary>
        /// <returns></returns>
        public MemoryStream FullScreenshot()
        {
            try
            {
                var stream = new MemoryStream();

                var bitmap = new Bitmap(_bounds.Width, _bounds.Height, PixelFormat.Format32bppArgb);

                using (var g = Graphics.FromImage(bitmap))
                {
                    g.CopyFromScreen(new Point(0, 0), new Point(0, 0), _bounds.Size);
                }

                bitmap = new Bitmap(bitmap, SIZE_WIDTH, SIZE_HEIGHT);

                bitmap.Save(stream, ImageFormat.Jpeg);

                return stream;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
