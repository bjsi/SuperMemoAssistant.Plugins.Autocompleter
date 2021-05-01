using KTrie;
using SuperMemoAssistant.Extensions;
using SuperMemoAssistant.Services;
using System.Collections.Generic;
using System.Linq;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public class SuggestionSource
  {
    private StringTrie<string> Words { get; }
    private AutocompleterCfg Config => Svc<AutocompleterPlugin>.Plugin.Config;

    public SuggestionSource()
    {
      Words = new StringTrie<string>();
    }

    public SuggestionSource(StringTrie<string> trie)
    {
      Words = trie;
    }

    public void AddWords(IEnumerable<StringEntry<string>> words)
    {
      foreach (var word in words)
      {
        Words[word.Key] = word.Value;
      }
    }

    public IEnumerable<string> GetWords(string prefix)
    {
      return Words.GetByPrefix(prefix)
        ?.Take(Config.MaxResults)
        ?.Select(x => x.Value)
        ?.Where(x => x != prefix);
    }

    public string GetValue(string key)
    {
      return Words[key];
    }

    public SuggestionSource Copy()
    {
      var ser = Words.Serialize();
      var des = ser.Deserialize<StringTrie<string>>();
      return new SuggestionSource(des);
    }
  }
}
