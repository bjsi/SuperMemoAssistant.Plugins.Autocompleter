#region License & Metadata

// The MIT License (MIT)
// 
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the "Software"),
// to deal in the Software without restriction, including without limitation
// the rights to use, copy, modify, merge, publish, distribute, sublicense,
// and/or sell copies of the Software, and to permit persons to whom the 
// Software is furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.
// 
// 
// Created On:   7/11/2020 1:53:51 PM
// Modified By:  james

#endregion




namespace SuperMemoAssistant.Plugins.Autocompleter
{
  using System;
  using System.Collections.Generic;
  using System.Diagnostics.CodeAnalysis;
  using System.Linq;
  using System.Windows;
  using mshtml;
  using SuperMemoAssistant.Extensions;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
  using SuperMemoAssistant.Interop.SuperMemo.Core;
  using SuperMemoAssistant.Services;
  using SuperMemoAssistant.Services.Sentry;
  using SuperMemoAssistant.Sys.Remoting;

  // ReSharper disable once UnusedMember.Global
  // ReSharper disable once ClassNeverInstantiated.Global
  [SuppressMessage("Microsoft.Naming", "CA1724:TypeNamesShouldNotMatchNamespaces")]
  public class AutocompleterPlugin : SentrySMAPluginBase<AutocompleterPlugin>
  {
    #region Constructors

    /// <inheritdoc />
    public AutocompleterPlugin() : base("Enter your Sentry.io api key (strongly recommended)") { }

    #endregion


    #region Properties Impl - Public

    /// <inheritdoc />
    public override string Name => "Autocompleter";

    /// <inheritdoc />
    public override bool HasSettings => true;
    public HTMLControlEvents Events { get; set; }
    public HashSet<string> Words = new HashSet<string>();
    public AutocompleterCfg Config;
    public IHTMLPopup CurrentPopup { get; set; }

    #endregion

    private void LoadConfig()
    {
      Config = Svc.Configuration.Load<AutocompleterCfg>() ?? new AutocompleterCfg();
    }

    #region Methods Impl

    /// <inheritdoc />
    protected override void PluginInit()
    {

      LoadConfig();

      SubscribeToHtmlEvents();

      Svc.SM.UI.ElementWdw.OnElementChanged += new ActionProxy<SMDisplayedElementChangedEventArgs>(ElementWdw_OnElementChanged);

    }

    private void ElementWdw_OnElementChanged(SMDisplayedElementChangedEventArgs obj)
    {

      FindWords();

    }

    private void FindWords()
    {

      Words = new HashSet<string>();

      var htmlCtrls = ContentUtils.GetHtmlCtrls();
      if (htmlCtrls.IsNull() || !htmlCtrls.Any())
        return;

      foreach (var htmlCtrl in htmlCtrls)
      {

        var doc = htmlCtrl?.GetDocument();
        var text = doc?.body?.innerText;
        if (text.IsNullOrEmpty())
          continue;

        Words.UnionWith(SplitIntoWords(text));

      }
    }

    // Run when last keypress was whitespace?
    private void UpdateWordsForFocusedHtmlCtrl()
    {

      int idx = Svc.SM.UI.ElementWdw.ControlGroup.FocusedControlIndex;
      var htmlCtrl = ContentUtils.GetFocusedHtmlDocument();
      if (idx < 0 || htmlCtrl.IsNull())
        return;

      Words.UnionWith(SplitIntoWords(htmlCtrl.body?.innerText));

    }

    private HashSet<string> SplitIntoWords(string text)
    {

      return text.IsNullOrEmpty()
        ? new HashSet<string>()
        : text
            ?.Split((char[])null)  // split on whitespace
            ?.Where(word => !word.IsNullOrEmpty())
            ?.Where(word => word.All(c => char.IsLetterOrDigit(c))) // filter invalid words
            ?.ToHashSet();

    }

    private void SubscribeToHtmlEvents()
    {

      var opts = new List<EventInitOptions>
      {
        new EventInitOptions(EventType.onkeydown),
        new EventInitOptions(EventType.onkeyup)
      };

      Events = new HTMLControlEvents(opts);

      Events.OnKeyDownEvent += Events_OnKeyDownEvent;
      Events.OnKeyUpEvent += Events_OnKeyUpEvent; // Debounce? check session info

    }

    private void Events_OnKeyUpEvent(object sender, IHTMLControlEventArgs e)
    {

      var ev = e.EventObj;

      int x = ev.clientX;
      int y = ev.clientY;

      var selObj = ContentUtils.GetSelectionObject();
      if (selObj.IsNull() || !selObj.text.IsNullOrEmpty())
        return;

      var htmlDoc = ContentUtils.GetFocusedHtmlDocument();
      if (htmlDoc.IsNull())
        return;

      bool foundWord = false;
      while (selObj.moveStart("character", -1) == -1)
      {

        if (selObj.text.IsNullOrEmpty())
          return;

        char first = selObj.text.First();
        if (char.IsWhiteSpace(first))
        {

          foundWord = true;
          selObj.moveStart("character", 1);
          break;

        }
        // Break if word contains punctuation
        else if (char.IsPunctuation(first))
          break;
      }

      if (!foundWord)
        return;

      IEnumerable<string> matches = FindMatchingWords(selObj);
      if (matches.IsNull() || !matches.Any())
        return;

      if (CurrentPopup.IsNull())
      {
        Application.Current.Dispatcher.BeginInvoke((Action)(() =>
        { 

          CurrentPopup = PopupEx.CreatePopup();
          CurrentPopup.UpdatePopup(matches);
          CurrentPopup.ShowPopup(matches.Count(), x, y, htmlDoc.body);

        }));
      }
      else
      {
        Application.Current.Dispatcher.BeginInvoke((Action)(() => 
        {

          CurrentPopup.UpdatePopup(matches);
          CurrentPopup.ShowPopup(matches.Count(), x, y, htmlDoc.body);

        }));
      }

    }

    // TODO: Add Config options for things like searching only within the current ctrl
    // TODO: Switch to trie if slow
    private IEnumerable<string> FindMatchingWords(IHTMLTxtRange selObj)
    {

      if (selObj.IsNull() || selObj.text.IsNullOrEmpty())
        return Enumerable.Empty<string>();

      return Words
        // TODO: Kept getting NRE so added ??
        ?.Where(word => word.Contains(selObj.text ?? string.Empty))
        ?.Where(word => word.Length >= 3)
        ?.OrderBy(x => x.Length)
        ?.Take(Config.MaxResults);
      
    }

    private void Events_OnKeyDownEvent(object sender, IHTMLControlEventArgs e)
    {


      if (CurrentPopup.IsNull() || !CurrentPopup.isOpen)
        return;

      var ev = e.EventObj;

      bool esc = ev.keyCode == 27;
      bool tab = ev.keyCode == 9;
      bool arrowUp = ev.keyCode == 38;
      bool arrowDown = ev.keyCode == 40;

      if (esc)
      {
        CurrentPopup.Hide();
        CurrentPopup = null;
      }

      // Convert to static methods
      else if (tab)
      {
        //SelectCurrentMenuItem(CurrentPopup);
      }
      else if (arrowDown)
      {
        //HighlightNextMenuItem(CurrentPopup);
      }
      else if (arrowUp)
      {
        //HighlightPrevMenuItem(CurrentPopup);
      }

    }

    /// <inheritdoc />
    public override void ShowSettings()
    {
    }

    #endregion

    #region Methods

    #endregion
  }
}
