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

    // TODO: Store GUITHREAD INFO here?

    /// <summary>
    /// Evaluates Cursor Position with respect to client screen.
    /// </summary>
    public static Point EvaluateCaretPosition()
    {
      var caretPosition = new Point();

      // Fetch GUITHREADINFO
      var guiInfo = GetCaretPosition();

      caretPosition.X = (int)guiInfo.rcCaret.Left; // + 25;
      caretPosition.Y = (int)guiInfo.rcCaret.Bottom; // + 25;

      //NativeMethods.ClientToScreen(guiInfo.hwndCaret, out caretPosition);
      return caretPosition;

    }

    private static GUITHREADINFO GetCaretPosition()
    {
      var guiInfo = new GUITHREADINFO();
      guiInfo.cbSize = (uint)Marshal.SizeOf(guiInfo);

      // Get GuiThreadInfo into guiInfo
      NativeMethods.GetGUIThreadInfo(0, out guiInfo);
      return guiInfo;
    }

  }
}

