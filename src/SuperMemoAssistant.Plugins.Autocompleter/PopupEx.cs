using mshtml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public static class PopupEx
  {

    public static void SelectFirstMenuItem(this IHTMLPopup popup)
    {
      try
      {

        if (popup.IsNull())
          return;

        var doc = popup.document as IHTMLDocument2;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

    }

    public static IHTMLPopup CreatePopup()
    {
      try
      {

        var wdw = ContentUtils.GetFocusedHtmlWindow() as IHTMLWindow4;
        var popup = wdw?.createPopup() as IHTMLPopup;

        // Styling
        var doc = popup?.document as IHTMLDocument2;
        if (!doc.IsNull())
          doc.body.style.border = "solid black 1px";

        return popup;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return null; 
    }
   
    // TODO: Requires dispatcher?
    public static void UpdatePopup(this IHTMLPopup popup, IEnumerable<string> matches)
    {

      try
      {

        if (popup.IsNull() || !matches.Any())
          return;

        var doc = popup.document as IHTMLDocument2;
        var body = doc?.body as IHTMLDOMNode;
        if (body.IsNull())
          return;

        // Clear the popup
        doc.body.innerHTML = "";

        // Add all matches to the menu
        foreach (var match in matches)
        {

          var menuItem = doc.createElement("<div>");
          menuItem.innerText = match;
          body.appendChild(((IHTMLDOMNode)menuItem));

        }

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }
    }

    public static void ShowPopup(this IHTMLPopup popup, int matches, int x, int y, IHTMLElement body)
    {
      try
      {

        int height = (matches * 17) + 5;
        popup.Show(x, y, 150, height, body);

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

    }
  }
}
