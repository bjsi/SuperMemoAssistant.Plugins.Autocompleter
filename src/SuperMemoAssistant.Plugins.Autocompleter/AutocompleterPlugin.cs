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
  using System.Drawing;
  using System.Linq;
  using System.Reactive.Linq;
  using System.Runtime.InteropServices;
  using System.Threading;
  using System.Threading.Tasks;
  using System.Windows;
  using System.Windows.Input;
  using Gma.DataStructures.StringSearch;
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
    public Trie<string> Words = new Trie<string>();
    public AutocompleterCfg Config;
    public IHTMLPopup CurrentPopup { get; set; }

    private static readonly Dictionary<int, int> IEFontSizeToPixels = new Dictionary<int, int>
    {
      { 1, 8 },
      { 2, 10 },
      { 3, 12 },
      { 4, 14 },
      { 5, 18 },
      { 6, 24 },
      { 7, 36 }
    };

    #endregion

    private void LoadConfig()
    {
      Config = Svc.Configuration.Load<AutocompleterCfg>() ?? new AutocompleterCfg();
    }

    private EventHandler<IHTMLControlEventArgs> CreateThrottledEventHandler(
    EventHandler<IHTMLControlEventArgs> handler,
    TimeSpan throttle)
    {
      bool throttling = false;
      return (s, e) =>
      {
        if (throttling) return;
        handler(s, e);
        throttling = true;

        // TODO: Add await?
        Task.Delay(throttle).ContinueWith(_ => throttling = false);
      };
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

      Words = new Trie<string>();

      var htmlCtrls = ContentUtils.GetHtmlCtrls();
      if (htmlCtrls.IsNull() || !htmlCtrls.Any())
        return;

      foreach (KeyValuePair<int, IControlHtml> kvpair in htmlCtrls)
      {

        int ctrlIdx = kvpair.Key;
        var htmlCtrl = kvpair.Value;

        var doc = htmlCtrl?.GetDocument();
        var text = doc?.body?.innerText;
        if (text.IsNullOrEmpty())
          continue;

        var words = SplitIntoWords(text);

        words.ForEach(word => Words.Add(word, word));

      }
    }

    // Run when last keypress was whitespace?
    private void UpdateWordsForFocusedHtmlCtrl()
    {

      int idx = Svc.SM.UI.ElementWdw.ControlGroup.FocusedControlIndex;
      var htmlCtrl = ContentUtils.GetFocusedHtmlDocument();
      if (idx < 0 || htmlCtrl.IsNull())
        return;

      //Words.UnionWith(SplitIntoWords(htmlCtrl.body?.innerText));

    }

    private HashSet<string> SplitIntoWords(string text)
    {

      return text.IsNullOrEmpty()
        ? new HashSet<string>()
        : text
            ?.Split((char[])null)  // split on whitespace
            ?.Where(word => !word.IsNullOrEmpty())
            ?.Where(word => word.Length > 3)
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

      // Throttles KeyUp to minimize the number of updates during continuous typing
      Events.OnKeyUpEvent += CreateThrottledEventHandler((s, e) => Events_OnKeyUpEvent(s, e), TimeSpan.FromSeconds(0.2));
    }

    public class LastPartialWord
    {
      public string Text { get; set; }
      public int Width { get; set; }
      public LastPartialWord(string word, int width)
      {
        this.Text = word;
        this.Width = width;
      }
    }

    private double GetWordWidth(string word, int fontSize, string fontName)
    {

      Font stringFont = new Font(fontName, fontSize);
      SizeF stringSize = new SizeF();
      stringSize = System.Windows.Forms.TextRenderer.MeasureText(word, stringFont);
      return stringSize.Width;

    }

    private LastPartialWord GetLastPartialWord(IHTMLTxtRange selObj)
    {

      LastPartialWord word = null;

      if (selObj.IsNull())
        return null;

      var duplicate = selObj.duplicate();
      while (duplicate.moveStart("character", -1) == -1)
      {
        if (duplicate.text.IsNullOrEmpty())
          return null;

        char first = duplicate.text.First();
        if (char.IsWhiteSpace(first))
        {

          duplicate.moveStart("character", 1);

          int fontsize = IEFontSizeToPixels[duplicate.QueryFontSize()];
          string fontname = duplicate.QueryFontName();
          int width = (int)GetWordWidth(duplicate.text, fontsize, fontname);
          word = new LastPartialWord(duplicate.text, width);
          break;

        }
        // Break if word contains punctuation
        else if (char.IsPunctuation(first))
          break;
      }

      return word;

    }

    private void Events_OnKeyUpEvent(object sender, IHTMLControlEventArgs e)
    {

      var ev = e.EventObj;

      var key = ev.keyCode;
      if (key == 27) // Esc
      {
        CurrentPopup?.Hide();
        return;
      }

      var selObj = ContentUtils.GetSelectionObject();
      if (selObj.IsNull() || !selObj.text.IsNullOrEmpty())
        return;

      var word = GetLastPartialWord(selObj);
      if (word.IsNull() || word.Text.IsNullOrEmpty())
      {
        CurrentPopup?.Hide();
        return;
      }

      IEnumerable<string> matches = FindMatchingWords(word.Text);
      if (matches.IsNull() || !matches.Any())
      {
        CurrentPopup?.Hide();
        return;
      }

      if (matches.Count() == 1 && matches.First() == word.Text)
      {
        CurrentPopup?.Hide();
        return;
      }

      var htmlDoc = ContentUtils.GetFocusedHtmlDocument();
      if (htmlDoc.IsNull())
        return;

      var caretPos = CaretPos.EvaluateCaretPosition();
      int x = caretPos.X - word.Width + 3;
      int y = caretPos.Y;

      Application.Current.Dispatcher.BeginInvoke((Action)(() =>
      {

        CurrentPopup = PopupEx.CreatePopup();
        CurrentPopup.UpdatePopup(matches);
        CurrentPopup.ShowPopup(matches.Count(), x, y, htmlDoc.body);

      }));

    }

    // TODO: Add Config options for things like searching only within the current ctrl?
    // TODO: case-insensitive search?
    private IEnumerable<string> FindMatchingWords(string word)
    {

      return word.IsNullOrEmpty()
        ? null
        : Words.Retrieve(word);

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
