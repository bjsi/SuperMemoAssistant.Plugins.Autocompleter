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
  using System.Reactive.Concurrency;
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
  using Newtonsoft.Json;
  using SuperMemoAssistant.Extensions;
  using SuperMemoAssistant.Interop.SuperMemo.Content.Controls;
  using SuperMemoAssistant.Interop.SuperMemo.Core;
  using SuperMemoAssistant.Plugins.Autocompleter.Interop;
  using SuperMemoAssistant.Services;
  using SuperMemoAssistant.Services.IO.HotKeys;
  using SuperMemoAssistant.Services.Sentry;
  using SuperMemoAssistant.Services.UI.Configuration;
  using SuperMemoAssistant.Sys.Remoting;
  using static SuperMemoAssistant.Plugins.Autocompleter.HtmlEventEx;


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
    
    /// <summary>
    /// Autocompleter service published and available for other plugins
    /// </summary>
    public AutocompleterSvc _autocompleterSvc = new AutocompleterSvc();

    /// <summary>
    /// Either points to the DefaultSuggestionSource, or a plugin can use the service
    /// to point at a custom source.
    /// </summary>
    private Trie<string> CurrentSuggestionSource;

    /// <summary>
    /// Optional converter dictionary used to convert accepted menu items.
    /// Useful for expanding snippets.
    /// </summary>
    public Dictionary<string, string> AcceptedSuggestionConverter { get; set; }

    /// <summary>
    /// Dictionary words + current element words.
    /// </summary>
    public Trie<string> DefaultSuggestionSource { get; set; }

    /// <summary>
    /// Name of the plugin that provided the suggestion words.
    /// </summary>
    public string SuggestionSourcePluginName { get; set; }

    /// <summary>
    /// Serialized Trie<string> of the top 10k English words by frequency > 3 letters:
    /// </summary>
    private string BaseSuggestionSource { get; set; }

    // TODO: Does this actually get used?

    /// <summary>
    /// Populated on ElementChanged Event
    /// </summary>
    public HashSet<string> InitialWordSet = new HashSet<string>();

    /// <summary>
    /// The current popup window instance.
    /// </summary>
    public HtmlPopup CurrentPopup { get; set; }

    public AutocompleterCfg Config;

    //
    // Html events
    private HtmlEvent _htmlKeydownEvent { get; set; }
    private HtmlEvent _htmlKeyupEvent { get; set; }

    /// <summary>
    /// Queue for event handler code to be run not on the UI thread
    /// </summary>
    private EventfulConcurrentQueue<Action> JobQueue = new EventfulConcurrentQueue<Action>();

    /// <summary>
    /// True after SMA / SM is closed and dispose is called.
    /// </summary>
    private bool HasExited = false;

    #endregion

    /// <inheritdoc />
    public override void ShowSettings()
    {
      ConfigurationWindow.ShowAndActivate(HotKeyManager.Instance, Config);
    }

    private void LoadConfig()
    {
      Config = Svc.Configuration.Load<AutocompleterCfg>() ?? new AutocompleterCfg();
    }


    //// TODO: Change to ReactiveEx
    //private EventHandler<IControlHtmlEventArgs> CreateThrottledEventHandler(
    //EventHandler<IControlHtmlEventArgs> handler,
    //TimeSpan throttle)
    //{
    //  bool throttling = false;
    //  return async (s, e) =>
    //  {
    //    if (throttling) return;
    //    handler(s, e);
    //    throttling = true;

    //    // TODO: Is it useful to have await here?
    //    await Task.Delay(throttle).ContinueWith(_ => throttling = false);
    //  };
    //}

    #region Methods Impl

    /// <inheritdoc />
    protected override void PluginInit()
    {

      LoadConfig();

      PublishService<IAutocompleterSvc, AutocompleterSvc>(_autocompleterSvc);

      BaseSuggestionSource = CreateBaseSuggestionSource();
      CurrentSuggestionSource = DefaultSuggestionSource;
      SuggestionSourcePluginName = Name;

      Svc.SM.UI.ElementWdw.OnElementChanged += new ActionProxy<SMDisplayedElementChangedEventArgs>(OnElementChanged);

      // Job Queue thread
      _ = Task.Factory.StartNew(HandleJobs, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);

    }

    private static string CreateBaseSuggestionSource()
    {

      var words = Words.English;
      if (words.IsNull() || !words.Any())
        return null;

      var trie = new Trie<string>();

      foreach (var word in words)
      {
        trie.Add(word, word);
      }

      return JsonConvert.SerializeObject(trie);

    }

    protected override void Dispose(bool disposing)
    {

      HasExited = true;
      base.Dispose(disposing);

    }

    private void HandleJobs()
    {
      while (!HasExited)
      {
        JobQueue.DataAvailableEvent.WaitOne(3000);
        while (JobQueue.TryDequeue(out var action))
        {
          try
          {
            action();
          }
          catch (RemotingException) { }
          catch (Exception e)
          {
            LogTo.Error($"Exception {e} caught in job queue thread");
          }
        }
      }
    }

    private void OnElementChanged(SMDisplayedElementChangedEventArgs obj)
    {

      FindWords();

      SuggestionSourcePluginName = Name;
      CurrentSuggestionSource = DefaultSuggestionSource;

      SubscribeToHtmlEvents();

    }

    private void SubscribeToKeydownEvent(IHTMLElement2 body)
    {

      if (body.IsNull())
        return;

      try
      {

        _htmlKeydownEvent = new HtmlEvent();
        body.SubscribeTo(EventType.onkeydown, _htmlKeydownEvent);
        _htmlKeydownEvent.OnEvent += _htmlKeydownEvent_OnEvent;

        Observable
          .FromEventPattern<IControlHtmlEventArgs>(
             h => _htmlKeydownEvent.OnEvent += h,
             h => _htmlKeydownEvent.OnEvent -= h
            )
          .Throttle(TimeSpan.FromMilliseconds(600))
          .Subscribe(_ => JobQueue.Enqueue(UpdateWordsForFocusedHtmlCtrl));

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }

    }

    private void SubscribeToKeyupEvent(IHTMLElement2 body)
    {

      if (body.IsNull())
        return;

      try
      {

        _htmlKeyupEvent = new HtmlEvent();
        body.SubscribeTo(EventType.onkeyup, _htmlKeyupEvent);
        _htmlKeyupEvent.OnEvent += _htmlKeyupEvent_OnEvent;

        // TODO: Use Reactivex!!!!

        // Throttles KeyUp to minimize the number of updates during continuous typing
        // TODO: Doesn't use ReactiveEx because accessing the eventObject throws errors
        // because of threading? 

        //_htmlKeyupEvent.OnEvent += CreateThrottledEventHandler(
        //  (s, e) => _htmlKeyupEvent_OnEvent(s, e),
        //            TimeSpan.FromSeconds(0.1));

        Observable
          .FromEventPattern<IControlHtmlEventArgs>(
            h => _htmlKeyupEvent.OnEvent += _htmlKeyupEvent_OnEvent,
            h => _htmlKeyupEvent.OnEvent -= _htmlKeyupEvent_OnEvent
          )
          .Throttle(TimeSpan.FromMilliseconds(200))
          // TODO: Test
          .ObserveOn(Scheduler.CurrentThread)
          .SubscribeOn(Scheduler.CurrentThread)
          .Subscribe(x => _htmlKeyupEvent_OnEvent(x.Sender, x.EventArgs));

      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }


    }

    /// <summary>
    /// Subscribe to keydown and keyup events
    /// </summary>
    private void SubscribeToHtmlEvents()
    {

      var htmlDoc = ContentUtils.GetFocusedHtmlDocument();
      var body = htmlDoc?.body as IHTMLElement2;
      if (body.IsNull())
        return;

      SubscribeToKeydownEvent(body);
      SubscribeToKeyupEvent(body);

    }

    private void FindWords()
    {

      InitialWordSet = new HashSet<string>();
      DefaultSuggestionSource = BaseSuggestionSource.Deserialize<Trie<string>>();

      var htmlCtrls = ContentUtils.GetHtmlCtrls();
      if (htmlCtrls.IsNull() || !htmlCtrls.Any())
        return;

      foreach (var htmlCtrl in htmlCtrls)
      {

        var doc = htmlCtrl?.GetDocument();
        var text = doc?.body?.innerText;
        if (text.IsNullOrEmpty())
          continue;

        var words = SplitIntoWords(text);
        if (words.IsNull() || !words.Any())
          return;

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
      if (words.IsNull() || !words.Any())
        return;

      words.ExceptWith(InitialWordSet);

      foreach (var word in words)
      {
        DefaultSuggestionSource.Add(word, word);
      }

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
            ?.Select(word => word.Trim(Words.PunctuationAndSymbols))
            ?.Where(word => word.All(c => char.IsLetterOrDigit(c))) // filter invalid words
            ?.ToHashSet();

    }

    private void _htmlKeydownEvent_OnEvent(object sender, IControlHtmlEventArgs e)
    {

      var ev = e?.EventObj;
      if (ev.IsNull())
        return;

      int key = ev.keyCode;

      bool esc = ev.keyCode == 27;
      bool arrowUp = ev.keyCode == 38;
      bool arrowDown = ev.keyCode == 40;
      bool arrowRight = ev.keyCode == 39;

      if (CurrentPopup.IsNull() || !CurrentPopup.IsOpen())
        return;

      if (!(arrowDown || arrowRight || arrowUp || esc))
        return;

      ev.returnValue = false;

      JobQueue.Enqueue(() => CurrentPopup.HandleKeydownInput(key));

    }

    private void _htmlKeyupEvent_OnEvent(object sender, IControlHtmlEventArgs e)
    {

      var ev = e?.EventObj;
      if (ev.IsNull())
        return;

      var key = ev.keyCode;
      int x = ev.screenX;
      int y = ev.screenY;

      // The action to execute depending on input

      Action action = null;

      // If there is a popup open already
      // pass input to the popup to handle

      if (!CurrentPopup.IsNull())
      {
        action = CurrentPopup.HandleKeyupInput(key, x, y);
      }

      // Else, handle input here
      else
      {

        switch (key)
        {

          case 27: // ESC
            action = () => CurrentPopup?.Hide();
            break;

          case 39: // Right Arrow
          case 38: // Up Arrow
          case 40: // Down Arrow 
            return;

          default:
            action = () => ShowNewAutocompleteWdw();
            break;

        }

      }

      if (!action.IsNull()) 
        JobQueue.Enqueue(action);

    }

    private void ShowNewAutocompleteWdw()
    {

      var selObj = ContentUtils.GetSelectionObject();
      if (selObj.IsNull() || !selObj.text.IsNullOrEmpty())
        return;

      var lastWord = ContentUtils.GetLastPartialWord(selObj);
      if (lastWord.IsNull() || lastWord.Text.IsNullOrEmpty())
      {
        CurrentPopup?.Hide();
        return;
      }

      var matches = FindMatchingWords(lastWord.Text);
      if (matches.IsNull() || !matches.Any())
      {
        CurrentPopup?.Hide();
        return;
      }

      if (matches.Count() == 1 && matches.First() == lastWord.Text)
      {
        CurrentPopup?.Hide();
        return;
      }

      var htmlDoc = ContentUtils.GetFocusedHtmlDocument();
      if (htmlDoc.IsNull())
        return;

      // Position the autocomplete window under the first letter of the last partial word
      var caretPos = CaretPos.EvaluateCaretPosition();
      int x = caretPos.X - lastWord.Width + 3;
      int y = caretPos.Y;
      int matchLength = lastWord.Text.Length;

      Application.Current.Dispatcher.BeginInvoke((Action)(() =>
      {

        var wdw = ContentUtils.GetFocusedHtmlWindow() as IHTMLWindow4;
        var htmlDoc = ((IHTMLWindow2)wdw)?.document;
        if (wdw.IsNull() || htmlDoc.IsNull())
          return;

        CurrentPopup = wdw.CreatePopup();
        CurrentPopup.AddContent(matches, matchLength);
        CurrentPopup.Show(x, y);

      }));
    }

    private IEnumerable<string> FindMatchingWords(string word)
    {

      return word.IsNullOrEmpty()
        ? null
        : CurrentSuggestionSource.Retrieve(word)?.Take(Config.MaxResults);

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
