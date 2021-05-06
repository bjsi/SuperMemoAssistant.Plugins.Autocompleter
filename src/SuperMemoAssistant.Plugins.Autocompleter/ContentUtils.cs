using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public static class ContentUtils
  {
    private static IHTMLTxtRange SelectLastPartialWord(IHTMLTxtRange selObj)
    {
      selObj.moveStart("word", -1);
      return selObj;
    }

    public static string GetLastPartialWord(IHTMLTxtRange selObj)
    {
      if (selObj == null)
        return null;

      var selected = SelectLastPartialWord(selObj);
      if (selected == null)
        return null;

      return selected.text;
    }

    /// <summary>
    /// Get the selection object representing the currently highlighted text in SM.
    /// </summary>
    /// <returns>IHTMLTxtRange object or null</returns>
    public static IHTMLTxtRange GetSelectionObject()
    {
      try
      {
        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
        var htmlDoc = htmlCtrl?.GetDocument();
        var sel = htmlDoc?.selection;

        if (!(sel?.createRange() is IHTMLTxtRange textSel))
          return null;

        return textSel;
      }
      catch (COMException) { }
      catch (UnauthorizedAccessException) { }

      return null;
    }

    /// <summary>
    /// Get the IHTMLDocument2 object representing the focused html control of the current element.
    /// </summary>
    /// <returns>IHTMLDocument2 object or null</returns>
    public static IHTMLDocument2 GetFocusedHtmlDocument()
    {
      try
      {
        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
        return htmlCtrl?.GetDocument();
      }
      catch (COMException) { }
      catch (UnauthorizedAccessException) { }

      return null;
    }

    public static List<IControlHtml> GetHtmlCtrls()
    {
      var ret = new List<IControlHtml>();

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      if (ctrlGroup == null)
        return ret;

      for (int i = 0; i < ctrlGroup.Count; i++)
      {
        var htmlCtrl = ctrlGroup[i].AsHtml();
        if (htmlCtrl != null)
          ret.Add(htmlCtrl);
      }

      return ret;
    }

    /// <summary>
    /// Get the IHTMLWindow2 object for the currently focused HtmlControl
    /// </summary>
    /// <returns>IHTMLWindow2 object or null</returns>
    public static IHTMLWindow2 GetFocusedHtmlWindow()
    {
      try
      {
        var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
        var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
        var htmlDoc = htmlCtrl?.GetDocument();
        if (htmlDoc == null)
          return null;

        return htmlDoc.parentWindow;
      }
      catch (COMException) { }
      catch (UnauthorizedAccessException) { }

      return null;
    }
  }
}
