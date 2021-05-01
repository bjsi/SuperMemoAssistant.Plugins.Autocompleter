using mshtml;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public class HtmlKeyInfo
  {

    public int keyCode { get; }
    public bool shiftKey { get; }
    public bool ctrlKey { get; }
    public bool altKey { get; }

    public HtmlKeyInfo(IHTMLEventObj ev)
    {
      this.keyCode = ev.keyCode;
      this.shiftKey = ev.shiftKey;
      this.ctrlKey = ev.ctrlKey;
      this.altKey = ev.altKey;
    }
  }
}
