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

    private static IHTMLTxtRange SelectLastPartialWord(IHTMLTxtRange selObj)
    {

      bool found = false;
      var duplicate = selObj.duplicate();
      while (duplicate.moveStart("character", -1) == -1)
      {
        if (duplicate.text.IsNullOrEmpty())
          return null;

        char first = duplicate.text.First();
        if (char.IsWhiteSpace(first))
        {
          duplicate.moveStart("character", 1);
          found = true;
        }
        // Break if word contains punctuation
        else if (char.IsPunctuation(first))
        {
          found = true;
          break;
        }
      }

      return found
        ? duplicate
        : null;

    }

    public static  LastPartialWord GetLastPartialWord(IHTMLTxtRange selObj)
    {

      if (selObj.IsNull())
        return null;

      var selected = SelectLastPartialWord(selObj);
      if (selected.IsNull())
        return null;

      return MeasureWord.CalculateLastPartialWord(selected);

    }

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

    public static List<IControlHtml> GetHtmlCtrls()
    {

      var ret = new List<IControlHtml>();

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      if (ctrlGroup.IsNull())
        return ret;

      for (int i = 0; i < ctrlGroup.Count; i++)
      {
        var htmlCtrl = ctrlGroup[i].AsHtml();
        if (!htmlCtrl.IsNull())
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

      var ctrlGroup = Svc.SM.UI.ElementWdw.ControlGroup;
      var htmlCtrl = ctrlGroup?.FocusedControl?.AsHtml();
      var htmlDoc = htmlCtrl?.GetDocument();
      if (htmlDoc == null)
        return null;

      return htmlDoc.parentWindow;

    }
  }
}
