using Anotar.Serilog;
using AutocompleterInterop;
using KTrie;
using mshtml;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Interop.SuperMemo.Core;
using SuperMemoAssistant.Services;
using SuperMemoAssistant.Services.IO.HotKeys;
using SuperMemoAssistant.Services.Sentry;
using SuperMemoAssistant.Services.UI.Configuration;
using SuperMemoAssistant.Sys.Remoting;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reactive.Concurrency;
using System.Reactive.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using static SuperMemoAssistant.Plugins.Autocompleter.HtmlEventEx;

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
// Created On:   4/29/2021 10:00:50 PM
// Modified By:  james

#endregion




namespace SuperMemoAssistant.Plugins.Autocompleter
{
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

    public AutocompleterSvc autocompleteSvc { get; } = new AutocompleterSvc();
    public SuggestionSource CurrentSuggestions { get; private set; } = new SuggestionSource();
    private SuggestionSource BaseSuggestions { get; set; } = new SuggestionSource();
    public HtmlPopup CurrentPopup { get; set; }
    public AutocompleterCfg Config { get; private set; }
    public Dictionary<string, string> Converter { get; set; } = new Dictionary<string, string>
    {
      ["nats ℕ"] = "ℕ",
      ["rats ℚ"] = "ℚ",
      ["reals ℝ"] = "ℝ",
      ["isin ∊"] = "∊",
      ["notin ∉"] = "∉",
      ["delta ∆"] = "∆",
      ["prod ×"] = "×",
      ["empty Ø"] = "Ø",
      ["ints ℤ"] = "ℤ",
    };

    public bool Enabled { get; set; } = true;

    //
    // Html events
    private HtmlEvent ContentChangedEvent { get; set; } = new HtmlEvent();
    private HtmlEvent KeyupEvent { get; set; } = new HtmlEvent();
    private HtmlEvent KeydownEvent { get; set; } = new HtmlEvent();
    #endregion


    #region Methods Impl

    public override void ShowSettings()
    {
      ConfigurationWindow.ShowAndActivate("Autocompleter", HotKeyManager.Instance, Config);
    }

    private void LoadConfig()
    {
      Config = Svc.Configuration.Load<AutocompleterCfg>() ?? new AutocompleterCfg();
    }

    /// <inheritdoc />
    protected override void OnSMStarted(bool wasSMAlreadyStarted)
    {
      LoadConfig();

      PublishService<IAutocompleterSvc, AutocompleterSvc>(autocompleteSvc);

      CreateInitialSuggestions();

      Svc.SM.UI.ElementWdw.OnElementChanged += new ActionProxy<SMDisplayedElementChangedEventArgs>(OnElementChanged);

      base.OnSMStarted(wasSMAlreadyStarted);
    }

    public void Disable()
    {
      Enabled = false;
    }

    public void Enable()
    {
      Enabled = true;
    }

    private void CreateInitialSuggestions()
    {
      var words = Words.LoadWords()?.Select(x => new StringEntry<string>(x, x));
      if (words != null)
      {
        BaseSuggestions.AddWords(words);
        BaseSuggestions.AddWords(new[]
        {
          new StringEntry<string>("nats ℕ", "nats ℕ"),
          new StringEntry<string>("rats ℚ", "rats ℚ"),
          new StringEntry<string>("reals ℝ", "reals ℝ"),
          new StringEntry<string>("isin ∊", "isin ∊"),
          new StringEntry<string>("notin ∉", "notin ∉"),
          new StringEntry<string>("delta ∆", "delta ∆"),
          new StringEntry<string>("prod ×", "prod ×"),
          new StringEntry<string>("empty Ø", "empty Ø"),
          new StringEntry<string>("ints ℤ", "ints ℤ"),
        });
        CurrentSuggestions = BaseSuggestions.Copy();
      }
    }

    public bool AddWords(IEnumerable<KeyValuePair<string, string>> keyValuePairs)
    {
      var words = keyValuePairs?.Select(x => new StringEntry<string>(x.Key, x.Value));
      if (words == null)
        return false;

      CurrentSuggestions.AddWords(words);
      BaseSuggestions.AddWords(words);
      return true;
    }

    private void OnElementChanged(SMDisplayedElementChangedEventArgs obj)
    {
      SubscribeToHtmlEvents();
      UpdateCurrentSuggestionSource();
    }

    private void SubscribeToContentChangedEvents(IHTMLElement2 body)
    {
      try
      {
        ContentChangedEvent = new HtmlEvent();
        body?.SubscribeTo(EventType.onkeydown, ContentChangedEvent);
        body?.SubscribeTo(EventType.onpaste, ContentChangedEvent);
        body?.SubscribeTo(EventType.oncut, ContentChangedEvent);

        Observable
          .FromEventPattern<IControlHtmlEventArgs>(
             h => ContentChangedEvent.OnEvent += h,
             h => ContentChangedEvent.OnEvent -= h
            )
          .Throttle(TimeSpan.FromSeconds(0.6))
          .SubscribeOn(TaskPoolScheduler.Default)
          .Subscribe(_ => UpdateCurrentSuggestionSource());
      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }
    }

    private void SubscribeToKeyupEvent(IHTMLElement2 body)
    {
      try
      {
        KeyupEvent = new HtmlEvent();
        body?.SubscribeTo(EventType.onkeyup, KeyupEvent);

        Observable
          .FromEventPattern<IControlHtmlEventArgs>(
            h => KeyupEvent.OnEvent += h,
            h => KeyupEvent.OnEvent -= h
          )
          .Select(x => new HtmlKeyInfo(x.EventArgs.EventObj))
          .Throttle(TimeSpan.FromSeconds(0.1))
          .SubscribeOn(TaskPoolScheduler.Default)
          .Subscribe(x => HandleKeyUp(x));
      }
      catch (UnauthorizedAccessException) { }
      catch (COMException) { }
    }

    /// <summary>
    /// Subscribe to keydown and keyup events
    /// </summary>

    private void SubscribeToHtmlEvents()
    {
      var htmlCtrls = ContentUtils.GetHtmlCtrls();
      if (htmlCtrls == null)
        return;

      foreach (var ctrl in htmlCtrls)
      {
        var doc = ctrl?.GetDocument();
        var body = doc?.body as IHTMLElement2;
        if (body == null)
          return;

        SubscribeToContentChangedEvents(body);
        SubscribeToKeyupEvent(body);
        SubscribeToKeydownEvent(body);
      }
    }

    private void SubscribeToKeydownEvent(IHTMLElement2 body)
    {
      KeydownEvent = new HtmlEvent();
      body?.SubscribeTo(EventType.onkeydown, KeydownEvent);
      KeydownEvent.OnEvent += KeydownEventHandler;
    }

    private void UpdateCurrentSuggestionSource()
    {
      LogTo.Debug("Update started.");
      var start = DateTime.Now;
      var htmlCtrls = ContentUtils.GetHtmlCtrls();
      if (htmlCtrls == null)
      {
        LogTo.Debug("Failed to update current suggestion source because htmlCtrls was null");
        return;
      }

      var newSuggestionss = BaseSuggestions.Copy();
      foreach (var htmlCtrl in htmlCtrls)
      {
        var doc = htmlCtrl?.GetDocument();
        var text = doc?.body?.innerText;
        if (string.IsNullOrEmpty(text))
          continue;

        var words = SplitIntoWords(text);
        if (words == null)
          return;

        newSuggestionss.AddWords(words);
      }
      CurrentSuggestions = newSuggestionss;
      LogTo.Debug("Update took: " + (DateTime.Now - start).TotalMilliseconds.ToString() + "ms");
    }

    private IEnumerable<StringEntry<string>> SplitIntoWords(string text)
    {
      return text
            ?.Split((char[])null, StringSplitOptions.RemoveEmptyEntries)
            ?.Where(word => word.Length > 3)
            ?.Select(word => word.Trim(Words.PunctuationAndSymbols))
            ?.Where(word => word.All(c => char.IsLetterOrDigit(c)))
            ?.Select(word => new StringEntry<string>(word, word));
    }

    private void KeydownEventHandler(object sender, IControlHtmlEventArgs e)
    {
      if (!Enabled) return;

      LogTo.Debug("Handling key down");
      var ev = e?.EventObj;
      if (ev == null)
        return;

      var key = new HtmlKeyInfo(ev);

      bool esc = key.keyCode == 27;
      bool arrowUp = key.keyCode == 38;
      bool arrowDown = key.keyCode == 40;
      bool arrowRight = key.keyCode == 39;

      if (CurrentPopup == null|| !CurrentPopup.IsOpen())
        return;

      if (!(arrowDown || arrowRight || arrowUp || esc))
        return;

      ev.returnValue = false;
      CurrentPopup.HandleKeydownInput(key);
    }

    private void HandleKeyUp(HtmlKeyInfo key)
    {
      if (!Enabled) return;
      // Right Arrow
      // Up Arrow
      // Down Arrow 
      if (new[] { 39, 38, 40 }.Contains(key.keyCode))
      {
        LogTo.Debug("KeyUp: Ignoring because it was an arrow key.");
        return;
      }
      else if (key.keyCode == 27)
      {
        LogTo.Debug("KeyUp: Escape pressed, closing popup");
        CurrentPopup?.Hide();
        return;
      }

      var selObj = ContentUtils.GetSelectionObject();
      if (selObj == null || !string.IsNullOrEmpty(selObj.text))
        return;

      var lastWord = ContentUtils.GetLastPartialWord(selObj);
      if (lastWord == null || string.IsNullOrEmpty(lastWord.Text))
      {
        CurrentPopup?.Hide();
        return;
      }

      var matches = CurrentSuggestions.GetWords(lastWord.Text);
      if (matches == null || !matches.Any())
      {
        CurrentPopup?.Hide();
        return;
      }

      if (matches.Count() == 1 && matches.First() == lastWord.Text)
      {
        CurrentPopup?.Hide();
        return;
      }

      if (CurrentPopup == null)
        ShowNewAutocompleteWdw(matches, lastWord.Text);
      else
        CurrentPopup?.Show(matches, lastWord.Text);
    }

    private void ShowNewAutocompleteWdw(IEnumerable<string> matches, string lastWord)
    {

      Application.Current.Dispatcher.BeginInvoke((Action)(() =>
      {

        var wdw = ContentUtils.GetFocusedHtmlWindow() as IHTMLWindow4;
        var htmlDoc = ((IHTMLWindow2)wdw)?.document;
        if (wdw == null || htmlDoc == null)
          return;

        LogTo.Debug("Creating and showing a new Autocomplete popup");
        CurrentPopup = wdw.CreatePopup();
        CurrentPopup.Show(matches, lastWord);
      }));
    }

    #endregion
  }
}
