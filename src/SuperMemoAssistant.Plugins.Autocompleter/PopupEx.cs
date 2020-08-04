using Anotar.Serilog;
using mshtml;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SuperMemoAssistant.Plugins.Autocompleter
{

  public class HtmlPopupEventArgs
  {

    public int x { get; set; }
    public int y { get; set; }
    public int w { get; set; }
    public int h { get; set; }
    public IHTMLPopup popup { get; set; }

    public HtmlPopupEventArgs(int x, int y, int w, int h, IHTMLPopup popup)
    {

      this.x = x;
      this.y = y;
      this.w = w;
      this.h = h;
      this.popup = popup;

    }

  }

  public class HtmlPopupOptions
  {

    public int x;
    public int y;
    public int MatchLength;
    public int NumberOfMatches;

    public HtmlPopupOptions(int x, int y, int matchLength, int numOfMatches)
    {

      this.x = x;
      this.y = y;
      this.MatchLength = matchLength;
      this.NumberOfMatches = numOfMatches;

    }
  }

  public static class PopupEx
  {
    public static HtmlPopup CreatePopup(this IHTMLWindow4 wdw)
    {
      return new HtmlPopup(wdw);
    }
  }

  public class HtmlPopup
  {

    private IHTMLPopup _popup { get; set; }

    // 
    // Events
    public event EventHandler<HtmlPopupEventArgs> OnShow;

    private Dictionary<string, string> AcceptedSuggestionConverter => Svc<AutocompleterPlugin>.Plugin.AcceptedSuggestionConverter;
    private AutocompleterSvc _autocompleterSvc => Svc<AutocompleterPlugin>.Plugin._autocompleterSvc;
    private string SuggestionSourcePluginName => Svc<AutocompleterPlugin>.Plugin.SuggestionSourcePluginName;

    /// <summary>
    /// Coordinates and sizing information
    /// </summary>
    private HtmlPopupOptions Options { get; set; }

    public HtmlPopup(IHTMLWindow4 wdw)
    {

      try
      {
        _popup = wdw?.createPopup() as IHTMLPopup;
        StylePopup();
      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    public Action HandleKeydownInput(int keyCode)
    {

      // Return an action based on the input

      Action action = null;

      switch (keyCode)
      {

        case 27: // ESC
          action = () => Hide();
          break;

        case 38: // Arrow Up
          action = () => SelectPreviousItem();
          break;

        case 40: // Arrow Down
          action = () => SelectNextItem();
          break;

        case 39: // Arrow Right
          action = () => InsertSelectedItem();
          break;

        default:

          // TODO: Show popup
          break;

      }

      return action;

    }

    private void InsertSelectedItem()
    {

      try
      {

        string word = null;

        if (_popup.IsNull())
          return;

        var selItem = GetSelectedItem();
        if (selItem.IsNull())
          return;

        var selObj = ContentUtils.GetSelectionObject();
        if (selObj.IsNull())
          return;

        // Replace the last partial word
        while (selObj.moveStart("character", -1) == -1)
        {

          if (selObj.text.IsNullOrEmpty())
            return;

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

        word = selItem.innerText;

        if (!AcceptedSuggestionConverter.IsNull())
        {
          if (!AcceptedSuggestionConverter.TryGetValue(word, out word))
          {
            LogTo.Warning($"Failed to find {word} in AcceptedSuggestionConverter dictionary");
            return;
          }
        }

        _autocompleterSvc?.InvokeSuggestionAccepted(word, SuggestionSourcePluginName);
        Hide();
        selObj.text = word;
        return;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    public Action HandleKeyupInput(int keyCode, int x, int y)
    {

      // Return an action depending on the input.

      Action action = null;

      switch (keyCode)
      {

        case 27: // ESC
          action = () => Hide();
          break;

        default:
          action = () => Show(x, y); // TODO
          break;

      }

      return action;

    }

    public bool IsOpen()
    {

      return _popup.IsNull()
        ? false
        : _popup.isOpen;

    }

    public void AddContent(IEnumerable<string> matches, int textLength)
    {

      try
      {

        if (_popup.IsNull() || !matches.Any())
          return;

        var doc = GetDocument();
        var body = doc?.body as IHTMLDOMNode;
        if (body.IsNull() || doc.IsNull())
          return;

        // Clear the popup content
        doc.body.innerHTML = "";

        // Add all matches to the menu
        foreach (var match in matches)
        {

          var menuItem = doc.createElement("<div>");
          menuItem.style.margin = "1px";
          menuItem.style.height = "17px";
          menuItem.style.border = "solid black 1px";

          menuItem.setAttribute("selected", 0);
          menuItem.innerHTML = "<span style='color: orange;'>" +
                                 "<B>" +
                                   match.Substring(0, textLength) +
                                 "</B>" +
                               "</span>" +
                               "<span style='color: grey'>" +
                                 match.Substring(textLength) +
                               "</span>";

          body.appendChild(((IHTMLDOMNode)menuItem));

        }

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    private void StylePopup()
    {

      var doc = GetDocument();
      if (!doc.IsNull())
      {

        doc.body.style.font = "Arial";
        doc.body.style.fontSize = "12px";
        doc.body.style.border = "solid black 1px";

      }

    }


    private IHTMLElement GetSelectedItem()
    {

      try
      {

        if (_popup.IsNull())
          return null;

        var doc = GetDocument();
        var children = doc?.body?.children as IHTMLElementCollection;
        var menuItems = children
          ?.Cast<IHTMLElement>()
          ?.Where(x => x.tagName.ToLower() == "div")
          ?.ToList();

        if (menuItems.IsNull() || menuItems.Count == 0)
          return null;

        IHTMLElement selected = null;

        for (int i = 0; i < menuItems.Count; i++)
        {
          var cur = menuItems[i];
          if (cur.tagName.ToLower() != "div")
            continue;

          if ((int)cur.getAttribute("selected") != 1)
            continue;

          selected = cur;
        }

        return selected;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return null;

    }

    private void SelectNextItem()
    {

      try
      {

        if (_popup.IsNull())
          return;

        var doc = GetDocument();
        var children = doc?.body?.children as IHTMLElementCollection;
        var childDivs = children
          ?.Cast<IHTMLElement>()
          ?.Where(x => x.tagName.ToLower() == "div")
          ?.ToList();

        int selIdx = -1;

        for (int i = 0; i < childDivs.Count; i++)
        {
          var cur = childDivs[i];

          if ((int)cur.getAttribute("selected") != 1)
            continue;

          selIdx = i;
          break;
        }

        if (selIdx == -1)
        {
          var prev = childDivs[0];
          SelectItem(prev);
          return;
        }
        else if (selIdx > -1)
        {
          var selected = childDivs[selIdx];
          UnselectItem(selected);

          var next = selIdx == childDivs.Count - 1
            ? childDivs[0]
            : childDivs[selIdx + 1];

          SelectItem(next);
        }
      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    private void SelectItem(IHTMLElement item)
    {

      if (item.IsNull())
        return;

      item.setAttribute("selected", 1);
      item.style.background = "lightblue";

    }

    private void UnselectItem(IHTMLElement item)
    {

      if (item.IsNull())
        return;

      try
      {

        item.setAttribute("selected", 0);
        item.style.background = "white";

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    private void SelectPreviousItem()
    {

      try
      {

        if (_popup.IsNull())
          return;

        var doc = GetDocument();
        var children = doc?.body?.children as IHTMLElementCollection;
        var childDivs = children
          ?.Cast<IHTMLElement>()
          ?.Where(x => x.tagName.ToLower() == "div")
          ?.ToList();

        if (childDivs.IsNull() || childDivs.Count == 0)
          return;

        int selIdx = -1;

        for (int i = 0; i < childDivs.Count; i++)
        {
          var cur = childDivs[i];

          if ((int)cur.getAttribute("selected") != 1)
            continue;

          selIdx = i;
          break;
        }

        if (selIdx == -1)
        {
          var next = childDivs[0];
          SelectItem(next);
          return;
        }
        else if (selIdx > -1)
        {

          var selected = childDivs[selIdx];
          UnselectItem(selected);

          var prev = selIdx == 0
            ? childDivs[childDivs.Count - 1]
            : childDivs[selIdx - 1];

          SelectItem(prev);

        }
      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }


    }

    private IEnumerable<string> GetItems()
    {

      try
      {

        var htmlDoc = GetDocument();
        var body = htmlDoc?.body;
        if (htmlDoc.IsNull() || body.IsNull())
          return null;
        
        var children = body?.children as IHTMLElementCollection;
        return children
          ?.Cast<IHTMLElement>()
          ?.Where(x => x.tagName.ToLower() == "div")
          ?.Select(x => x.innerText);

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

      return null;

    }

    public void Show(int x, int y)
    {

      try
      {

        if (_popup.IsNull())
          return;

        var items = GetItems();
        if (items.IsNull() || !items.Any())
          return;

        var coords = CalculateDimensions(items);
        if (!coords.HasValue)
          return;

        int width = (int)coords.Value.X;
        int height = (int)coords.Value.Y;

        var selObj = ContentUtils.GetSelectionObject();
        if (selObj.IsNull())
          return;

        var lastWord = ContentUtils.GetLastPartialWord(selObj);
        if (lastWord.IsNull() || lastWord.Text.IsNullOrEmpty())
        {
          Hide();
          return;
        }

        // Position the autocomplete window under the first letter of the last partial word
        var caretPos = CaretPos.EvaluateCaretPosition();
        x = caretPos.X - lastWord.Width + 3;
        y = caretPos.Y;
        int matchLength = lastWord.Text.Length;

        OnShow?.Invoke(this, new HtmlPopupEventArgs(x, y, width, height, _popup));
        _popup.Show(x, y, width, height, null);

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }


    }

    private Point? CalculateDimensions(IEnumerable<string> items)
    {

      if (items.IsNull() || !items.Any())
        return null;

      string longestWord = items
        .OrderByDescending(x => x.Length)
        .First();

      var coords = new Point();
      coords.X = MeasureWord.GetWordWidth(longestWord, 12, "Arial");
      if (coords.X == -1)
        return null;

      coords.Y = items.Count() * 19; // TODO

      return coords;

    }

    /// <summary>
    /// Hide the popup.
    /// </summary>
    public void Hide()
    {

      try
      {

        _popup?.Hide();

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    public IHTMLDocument2 GetDocument()
    {

      try
      {

        return _popup?.document as IHTMLDocument2;

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

      return null;

    }
  }
}
