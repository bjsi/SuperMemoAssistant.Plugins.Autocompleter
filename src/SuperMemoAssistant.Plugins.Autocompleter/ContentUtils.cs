using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public static class ContentUtils
  {
    /// <summary>
    /// Get the selection object representing the currently highlighted text in SM.
    /// </summary>
    /// <returns>IHTMLTxtRange object or null</returns>
    public static IHTMLTxtRange GetSelectionObject()
    {

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
      var htmlDoc = htmlCtrl?.GetDocument();
      var sel = htmlDoc?.selection;

      if (!(sel?.createRange() is IHTMLTxtRange textSel))
        return null;

      return textSel;

    }

    /// <summary>
    /// Get the IHTMLDocument2 object representing the focused html control of the current element.
    /// </summary>
    /// <returns>IHTMLDocument2 object or null</returns>
    public static IHTMLDocument2 GetFocusedHtmlDocument()
    {

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
      return htmlCtrl?.GetDocument();

    }

    public static Dictionary<int, IControlHtml> GetHtmlCtrls()
    {

      var ret = new Dictionary<int, IControlHtml>();

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      if (ctrlGroup.IsNull())
        return ret;

      for (int i = 0; i < ctrlGroup.Count; i++)
      {
        var htmlCtrl = ctrlGroup[i].AsHtml();
        if (!htmlCtrl.IsNull())
          ret.Add(i, htmlCtrl);
      }

      return ret;

    }

    /// <summary>
    /// Get the IHTMLWindow2 object for the currently focused HtmlControl
    /// </summary>
    /// <returns>IHTMLWindow2 object or null</returns>
    public static IHTMLWindow2 GetFocusedHtmlWindow()
    {

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
      var htmlDoc = htmlCtrl?.GetDocument();
      if (htmlDoc == null)
        return null;

      return htmlDoc.parentWindow;

    }
  }
}
