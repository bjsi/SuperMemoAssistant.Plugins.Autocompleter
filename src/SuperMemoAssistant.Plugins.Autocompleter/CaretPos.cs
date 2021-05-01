using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Services;
using System;
using System.Runtime.InteropServices;
using System.Windows;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public static class CaretPos
  {

    [DllImport("user32.dll")]
    static extern bool GetCaretPos(out Point lpPoint);

    /*- Converts window specific point to screen specific -*/
    [DllImport("user32.dll")]
    public static extern bool ClientToScreen(IntPtr hWnd, out Point position);

    [DllImport("user32.dll")]
    public static extern bool ScreenToClient(IntPtr hWnd, out Point position);

    /// <summary>
    /// Evaluates Cursor Position
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

