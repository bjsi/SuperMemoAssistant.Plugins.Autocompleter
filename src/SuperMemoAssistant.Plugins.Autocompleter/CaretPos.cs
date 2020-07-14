using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public static class CaretPos
  {

    /// <summary>
    /// Evaluates Cursor Position with respect to client screen.
    /// </summary>
    public static Point EvaluateCaretPosition()
    {
      var caretPosition = new Point();

      var guiInfo = GetCaretPosition();

      caretPosition.X = (int)guiInfo.rcCaret.Left;
      caretPosition.Y = (int)guiInfo.rcCaret.Bottom;

      return caretPosition;

    }

    private static GUITHREADINFO GetCaretPosition()
    {
      var guiInfo = new GUITHREADINFO();
      guiInfo.cbSize = (uint)Marshal.SizeOf(guiInfo);

      NativeMethods.GetGUIThreadInfo(0, out guiInfo);
      return guiInfo;
    }

  }
}

