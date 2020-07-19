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
  using System.Runtime.Remoting;
  using System.Threading;
  using System.Threading.Tasks;
  using System.Windows;
  using System.Windows.Input;
  using Anotar.Serilog;
  using Gma.DataStructures.StringSearch;
  using mshtml;
  using SuperMemoAssistant.Extensions;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
  using SuperMemoAssistant.Interop.SuperMemo.Core;
  using SuperMemoAssistant.Plugins.Autocompleter.Interop;
  using SuperMemoAssistant.Services;
  using SuperMemoAssistant.Services.IO.HotKeys;
  using SuperMemoAssistant.Services.Sentry;
  using SuperMemoAssistant.Services.UI.Configuration;
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


    private AutocompleterSvc _autocompleterSvc = new AutocompleterSvc();

    /// <summary>
    /// Either points to the DefaultSuggestionSource, or a plugin can use the service
    /// to point at a custom source.
    /// </summary>
    private Trie<string> CurrentSuggestionSource { get; set; }

    /// <summary>
    /// Optional converter dictionary used to convert accepted menu items.
    /// </summary>
    private Dictionary<string, string> AcceptedSuggestionConverter { get; set; }

    /// <summary>
    /// Dictionary words + current element words.
    /// </summary>
    public Trie<string> DefaultSuggestionSource = new Trie<string>();

    /// <summary>
    /// Name of the plugin that provided the suggestion words.
    /// </summary>
    private string SuggestionSourcePluginName { get; set; }

    /// <summary>
    /// Populated on ElementChanged Event
    /// </summary>
    public HashSet<string> InitialWordSet = new HashSet<string>();

    public IHTMLPopup CurrentPopup { get; set; }

    public AutocompleterCfg Config;

    private static readonly char[] PunctuationAndSymbols = new char[]
    {
      '.', '!', '?', ')', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '+', '=', '\\', '/', '<', '>', ','
    };

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

    /// <summary>
    /// Queue for event handler code to be run not on the UI thread
    /// </summary>
    private EventfulConcurrentQueue<Action> EventQueue = new EventfulConcurrentQueue<Action>();

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
      return async (s, e) =>
      {
        if (throttling) return;
        handler(s, e);
        throttling = true;

        // TODO: Is it useful to have await here?
        await Task.Delay(throttle).ContinueWith(_ => throttling = false);
      };
    }

    #region Methods Impl

    /// <inheritdoc />
    protected override void PluginInit()
    {

      LoadConfig();

      SubscribeToHtmlEvents();

      PublishService<IAutocompleterSvc, AutocompleterSvc>(_autocompleterSvc);

      CurrentSuggestionSource = DefaultSuggestionSource;
      SuggestionSourcePluginName = Name;

      Svc.SM.UI.ElementWdw.OnElementChanged += new ActionProxy<SMDisplayedElementChangedEventArgs>(ElementWdw_OnElementChanged);

      _ = Task.Factory.StartNew(DispatchEvents, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

    }

    private void DispatchEvents()
    {
      while (true)
      {
        EventQueue.DataAvailableEvent.WaitOne(3000);
        while (EventQueue.TryDequeue(out var action))
        {
          action();
        }
      }
    }

    private void ElementWdw_OnElementChanged(SMDisplayedElementChangedEventArgs obj)
    {

      InitialWordSet = new HashSet<string>();
      FindWords();

    }

    /// <inheritdoc />
    public override void ShowSettings()
    {
      ConfigurationWindow.ShowAndActivate(HotKeyManager.Instance, Config);
    }

    private void FindWords()
    {

      DefaultSuggestionSource = new Trie<string>();

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
        foreach (var word in words)
        {
          InitialWordSet.Add(word);
          DefaultSuggestionSource.Add(word, word);
        }

      }
    }

    // Run when last keypress was whitespace?
    private void UpdateWordsForFocusedHtmlCtrl()
    {

      var htmlCtrl = ContentUtils.GetFocusedHtmlDocument();
      var text = htmlCtrl?.body?.innerText;
      if (text.IsNullOrEmpty())
        return;

      var words = SplitIntoWords(text);
      words.ExceptWith(InitialWordSet);
      words.ForEach(w => DefaultSuggestionSource.Add(w, w));
      InitialWordSet.UnionWith(words);

    }

    private HashSet<string> SplitIntoWords(string text)
    {

      return text.IsNullOrEmpty()
        ? new HashSet<string>()
        : text
            ?.Split((char[])null)  // split on whitespace
            ?.Where(word => !word.IsNullOrEmpty())
            ?.Where(word => word.Length > 3)
            ?.Select(word => word.Trim(PunctuationAndSymbols))
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

      Observable
        .FromEventPattern<IHTMLControlEventArgs>(
           h => Events.OnKeyDownEvent += h,
           h => Events.OnKeyDownEvent -= h
          )
        .Throttle(TimeSpan.FromMilliseconds(600))
        .Subscribe(_ => EventQueue.Enqueue(UpdateWordsForFocusedHtmlCtrl));

      // Throttles KeyUp to minimize the number of updates during continuous typing
      // TODO: Doesn't use ReactiveEx because accessing the eventObject throws errors
      // because of threading? 

      Events.OnKeyUpEvent += CreateThrottledEventHandler(
        (s, e) => Events_OnKeyUpEvent(s, e),
                  TimeSpan.FromSeconds(0.1));
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

      if (word.IsNullOrEmpty() || fontSize < 0 || fontName.IsNullOrEmpty())
        return -1;

      Font stringFont = new Font(fontName, fontSize);
      var stringSize = System.Windows.Forms.TextRenderer.MeasureText(word, stringFont);
      stringFont.Dispose();
      return stringSize.IsNull()
        ? -1
        : stringSize.Width;

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

          var fontSizeIE = duplicate.QueryFontSize();
          int fontSizePx = IEFontSizeToPixels[fontSizeIE];
          string fontname = duplicate.QueryFontName();
          int width = (int)GetWordWidth(duplicate.text, fontSizePx, fontname);
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
      else if (key == 39) // Right Arrow
        return;
      else if (key == 38) // Up Arrow
        return;
      else if (key == 40) // Down Arrow 
        return;

      EventQueue.Enqueue(() =>
      {

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
        int matchLength = word.Text.Length;

        Application.Current.Dispatcher.BeginInvoke((Action)(() =>
        {

          CurrentPopup = PopupEx.CreatePopup();
          CurrentPopup.UpdatePopup(matches, matchLength);
          CurrentPopup.ShowPopup(matches.Count(), x, y, htmlDoc.body);

        }));

      });

    }

    private IEnumerable<string> FindMatchingWords(string word)
    {

      return word.IsNullOrEmpty()
        ? null
        : DefaultSuggestionSource.Retrieve(word)?.Take(Config.MaxResults);

    }

    public bool InsertCurrentSelection(IHTMLPopup popup, out string word)
    {

      word = null;

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
        
        word = selected.innerText;

        if (!AcceptedSuggestionConverter.IsNull())
        {
          if (!AcceptedSuggestionConverter.TryGetValue(word, out word))
          {
            LogTo.Warning($"Failed to find {word} in AcceptedSuggestionConverter dictionary");
            return false;
          }
        }

        selObj.text = word;
        return true;

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return false;

    }

    private void Events_OnKeyDownEvent(object sender, IHTMLControlEventArgs e)
    {

      var ev = e.EventObj;

      bool esc = ev.keyCode == 27;
      bool tab = ev.keyCode == 9; // Doesn't work

      bool arrowUp = ev.keyCode == 38;
      bool arrowDown = ev.keyCode == 40;
      bool arrowRight = ev.keyCode == 39;

      if (CurrentPopup.IsNull() || !CurrentPopup.isOpen)
        return;

      if (!arrowDown && !arrowRight && !arrowUp && !tab)
        return;

      ev.returnValue = false;

      EventQueue.Enqueue(() =>
      {

        if (esc)
          CurrentPopup?.Hide();

        else if (arrowRight)
        {

          if (InsertCurrentSelection(CurrentPopup, out var word))
          {
            _autocompleterSvc?.InvokeSuggestionAccepted(word, SuggestionSourcePluginName);
            CurrentPopup?.Hide();
          }

        }
        else if (arrowDown)
          CurrentPopup?.SelectNextMenuItem();

        else if (arrowUp)
          CurrentPopup?.SelectPrevMenuItem();

      });

    }


    public bool SetWordSuggestionSource(string pluginName, Trie<string> TrieOfWords)
    {

      if (TrieOfWords.IsNull())
        return false;

      CurrentSuggestionSource = TrieOfWords;
      SuggestionSourcePluginName = pluginName;
      return true;

    }

    public bool SetWordSuggestionSource(string pluginName, Trie<string> TrieOfWords, Dictionary<string, string> Converter)
    {

      if (TrieOfWords.IsNull() || Converter.IsNull())
        return false;

      CurrentSuggestionSource = TrieOfWords;
      AcceptedSuggestionConverter = Converter;
      SuggestionSourcePluginName = pluginName;
      return true;

    }

    public void ResetWordSuggestionSource()
    {

      CurrentSuggestionSource = DefaultSuggestionSource;
      SuggestionSourcePluginName = Name;
      AcceptedSuggestionConverter = null;

    }

    #endregion

    #region Methods

    #endregion
  }
}
