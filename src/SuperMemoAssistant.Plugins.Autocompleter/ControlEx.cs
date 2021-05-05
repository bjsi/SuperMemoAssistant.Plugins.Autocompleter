using SuperMemoAssistant.Interop.SuperMemo.UI.Element;
using SuperMemoAssistant.Sys.IO.Devices;
using System;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public static class ControlEx
  {
    public static bool SendKeys(this IElementWdw wdw,
                                HotKey hk,
                                int timeout = 100)
    {
      return SendKeys(wdw.Handle, hk, timeout);
    }

    public static bool SendKeys(IntPtr handle,
                                 HotKey hotKey,
                                 int timeout = 100)
    {
      if (handle.ToInt32() == 0)
        return false;

      if (hotKey.Alt && hotKey.Ctrl == false && hotKey.Win == false)
        return Keyboard.PostKeysAsync(
          handle,
          hotKey
        ).Wait(timeout);

      return Keyboard.PostKeysAsync(
        handle,
        hotKey
      ).Wait(timeout);
    }
  }
}
