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

    public bool SetWordSuggestionSource(Trie<string> TrieOfWords)
    {

      if (TrieOfWords.IsNull())
        return false;

      return Svc<AutocompleterPlugin>.Plugin.SetWordSuggestionSource(TrieOfWords);

    }

    public void ResetWordSuggestionSource()
    {

      Svc<AutocompleterPlugin>.Plugin.ResetWordSuggestionSource();

    }

    public event Action<SuggestionAcceptedEventArgs> OnSuggestionAccepted;

    public void InvokeSuggestionAccepted(string word)
    {
      
      if (word.IsNull())
        return;

      OnSuggestionAccepted?.Invoke(new SuggestionAcceptedEventArgs(word));

    }

  }
}
