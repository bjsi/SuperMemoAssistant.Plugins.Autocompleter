﻿using Anotar.Serilog;
using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
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
    public bool InsertSelectedItem()
    {
      try
      {
        string word = null;

        if (_popup == null)
          return false;

        var selItem = GetSelectedItem();
        if (selItem == null)
          return false;

        var selObj = ContentUtils.GetSelectionObject();
        if (selObj == null)
          return false;

        // Replace the last partial word
        while (selObj.moveStart("character", -1) == -1)
        {

          if (string.IsNullOrEmpty(selObj.text))
            return false;

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
        if (Svc<AutocompleterPlugin>.Plugin.Converter.ContainsKey(word))
          word = Svc<AutocompleterPlugin>.Plugin.Converter[word];


        Hide();
        selObj.text = word;
        if (word.Contains("<++>"))
        {
          var sel = ContentUtils.GetSelectionObject();
          sel.moveEnd("character", -word.Length);
          sel.select();
        }

        return true;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

      return false;

    }

    public bool IsOpen()
    {
      return _popup == null
        ? false
        : _popup.isOpen;
    }

    public void AddContent(IEnumerable<string> matches, int textLength)
    {

      try
      {

        if (_popup == null || !matches.Any())
          return;

        var doc = GetDocument();
        var body = doc?.body as IHTMLDOMNode;
        if (body == null || doc == null)
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
      if (doc != null)
      {
        //doc.body.style.font = "Arial";
        doc.body.style.fontSize = "12px";
        doc.body.style.border = "solid black 1px";
      }
    }

    public IHTMLElement GetSelectedItem()
    {

      try
      {

        if (_popup == null)
          return null;

        var doc = GetDocument();
        var children = doc?.body?.children as IHTMLElementCollection;
        var menuItems = children
          ?.Cast<IHTMLElement>()
          ?.Where(x => x.tagName.ToLower() == "div")
          ?.ToList();

        if (menuItems == null || menuItems.Count == 0)
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

    public void SelectNextItem()
    {

      try
      {

        if (_popup == null)
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

      if (item == null)
        return;

      item.setAttribute("selected", 1);
      item.style.background = "lightblue";

    }

    private void UnselectItem(IHTMLElement item)
    {

      if (item == null)
        return;

      try
      {

        item.setAttribute("selected", 0);
        item.style.background = "white";

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    public void SelectPreviousItem()
    {

      try
      {

        if (_popup == null)
          return;

        var doc = GetDocument();
        var children = doc?.body?.children as IHTMLElementCollection;
        var childDivs = children
          ?.Cast<IHTMLElement>()
          ?.Where(x => x.tagName.ToLower() == "div")
          ?.ToList();

        if (childDivs == null || childDivs.Count == 0)
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
        if (htmlDoc == null || body == null)
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

    public void Show(IEnumerable<string> matches, string lastWord)
    {
      try
      {
        if (_popup == null)
          return;

        AddContent(matches, lastWord.Length);

        var items = GetItems();
        if (items == null || !items.Any())
          return;

        var coords = CalculateDimensions(items);
        if (!coords.HasValue)
          return;

        int width = (int)coords.Value.X;
        int height = (int)coords.Value.Y;


        // Position the autocomplete window under the first letter of the last partial word
        var caretPos = CaretPos.EvaluateCaretPosition();
        var x = caretPos.X;
        var y = caretPos.Y;
        int matchLength = lastWord.Length;

        var ctrl = Svc.SM.UI.ElementWdw.ControlGroup.FocusedControl as IControlHtml;
        var body = ctrl.GetDocument()?.body;
        _popup.Show((int)x, (int)y, width, height, body);

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    private Point? CalculateDimensions(IEnumerable<string> items)
    {

      if (items == null || !items.Any())
        return null;

      string longestWord = items
        .OrderByDescending(x => x.Length)
        .FirstOrDefault();

      if (longestWord == null)
        return null;

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
