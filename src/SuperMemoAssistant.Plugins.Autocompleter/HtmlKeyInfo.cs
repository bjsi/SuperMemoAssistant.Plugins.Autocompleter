using mshtml;
using SuperMemoAssistant.Sys.IO.Devices;
using System.Windows.Forms;
using System.Windows.Input;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public class HtmlKeyInfo
  {

    public Keys Key { get; }
    public KeyModifiers Modifiers { get; } = KeyModifiers.None;

    public HtmlKeyInfo(IHTMLEventObj ev)
    {
      this.Key = ((Keys)ev.keyCode) & ~(Keys.Control | Keys.Shift | Keys.Alt | Keys.ControlKey | Keys.ShiftKey);
      if (ev.shiftKey)
        Modifiers |= KeyModifiers.Shift;
      if (ev.ctrlKey)
        Modifiers |= KeyModifiers.Ctrl;
      if (ev.altKey)
        Modifiers |= KeyModifiers.Alt;
    }
  }
}
