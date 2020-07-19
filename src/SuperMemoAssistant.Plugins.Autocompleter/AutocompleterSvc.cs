using Gma.DataStructures.StringSearch;
using PluginManager.Interop.Sys;
using SuperMemoAssistant.Plugins.Autocompleter.Interop;
using SuperMemoAssistant.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public class AutocompleterSvc : PerpetualMarshalByRefObject, IAutocompleterSvc
  {

    public bool SetWordSuggestionSource(string pluginName, Trie<string> TrieOfWords)
    {

      if (TrieOfWords.IsNull())
        return false;

      return Svc<AutocompleterPlugin>.Plugin.SetWordSuggestionSource(pluginName, TrieOfWords);

    }

    public bool SetWordSuggestionSource(string pluginName, Trie<string> TrieOfWords, Dictionary<string, string> Converter)
    {

      if (TrieOfWords.IsNull() || Converter.IsNull())
        return false;

      return Svc<AutocompleterPlugin>.Plugin.SetWordSuggestionSource(pluginName, TrieOfWords, Converter);

    }

    public void ResetWordSuggestionSource()
    {

      Svc<AutocompleterPlugin>.Plugin.ResetWordSuggestionSource();

    }

    public event Action<SuggestionAcceptedEventArgs> OnSuggestionAccepted;

    public void InvokeSuggestionAccepted(string word, string pluginName)
    {
      
      if (word.IsNullOrEmpty() || pluginName.IsNullOrEmpty())
        return;

      OnSuggestionAccepted?.Invoke(new SuggestionAcceptedEventArgs(word, pluginName));

    }

  }
}
