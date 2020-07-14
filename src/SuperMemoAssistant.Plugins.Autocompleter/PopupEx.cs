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

    public static IHTMLElement GetSelectedMenuItem(this IHTMLPopup popup)
    {

      try
      {

        if (popup.IsNull())
          return null;

        var doc = popup.document as IHTMLDocument2;
        var body = doc?.body?.children as IHTMLElementCollection;
        var children = body.Cast<IHTMLElement>()?.ToList();

        IHTMLElement selected = null;

        for (int i = 0; i < children.Count; i++)
        {
          var cur = children[i];
          if (cur.tagName.ToLower() != "div")
            continue;

          if ((int)cur.getAttribute("selected") != 1)
            continue;

          selected = children[i];
        }

        return selected;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return null;

    }

    public static bool InsertCurrentSelection(this IHTMLPopup popup)
    {

      try
      {

        if (popup.IsNull())
          return false;

        var selected = popup.GetSelectedMenuItem();
        if (selected.IsNull())
          return false;

        var selObj = ContentUtils.GetSelectionObject();
        if (selObj.IsNull())
          return false;

        // Replace the last partial word
        while (selObj.moveStart("character", -1) == -1)
        {

          char first = selObj.text.First();
          if (char.IsWhiteSpace(first))
          {
            selObj.moveStart("character", 1);
            break;
          }
          // Break if word contains punctuation
          else if (char.IsPunctuation(first))
            break;
        }

        selObj.text = selected.innerText;
        return true;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return false;

    }

    public static void SelectPrevMenuItem(this IHTMLPopup popup)
    {

      try
      {

        if (popup.IsNull())
          return;

        var doc = popup.document as IHTMLDocument2;
        var body = doc?.body?.children as IHTMLElementCollection;
        var children = body.Cast<IHTMLElement>()?.ToList();
        int selIdx = -1;

        for (int i = 0; i < children.Count; i++)
        {
          var cur = children[i];
          if (cur.tagName.ToLower() != "div")
            continue;

          if ((int)cur.getAttribute("selected") != 1)
            continue;

          selIdx = i;
          break;
        }

        if (selIdx == -1)
        {
          var next = children[0];
          next.SelectMenuItem();
          return;
        }
        else if (selIdx > -1)
        {

          var selected = children[selIdx];
          selected.UnselectMenuItem();

          var prev = selIdx == 0
            ? children[children.Count - 1]
            : children[selIdx - 1];

          prev.SelectMenuItem();

        }
      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

    }

    public static void SelectNextMenuItem(this IHTMLPopup popup)
    {
      try
      {

        if (popup.IsNull())
          return;

        var doc = popup.document as IHTMLDocument2;
        var body = doc?.body?.children as IHTMLElementCollection;
        var children = body.Cast<IHTMLElement>()?.ToList();
        int selIdx = -1;

        for (int i = 0; i < children.Count; i++)
        {
          var cur = children[i];
          if (cur.tagName.ToLower() != "div")
            continue;

          if ((int)cur.getAttribute("selected") != 1)
            continue;

          selIdx = i;
          break;
        }

        if (selIdx == -1)
        {
          var prev = children[0];
          prev.SelectMenuItem();
          return;
        }
        else if (selIdx > -1)
        {
          var selected = children[selIdx];
          selected.UnselectMenuItem();

          var next = selIdx == children.Count - 1
            ? children[0]
            : children[selIdx + 1];

          next.SelectMenuItem();
        }
      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }
    }

    private static void UnselectMenuItem(this IHTMLElement el)
    {

      if (el.IsNull())
        return;

      el.setAttribute("selected", 0);
      el.style.background = "white";

    }

    private static void SelectMenuItem(this IHTMLElement el)
    {

      if (el.IsNull())
        return;

      el.setAttribute("selected", 1);
      el.style.background = "lightblue";

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
          menuItem.setAttribute("selected", 0);
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
