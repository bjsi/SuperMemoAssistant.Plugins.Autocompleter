using AutocompleterInterop;
using PluginManager.Interop.Sys;
using SuperMemoAssistant.Services;
using System.Collections.Generic;
using System.Linq;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public class AutocompleterSvc : PerpetualMarshalByRefObject, IAutocompleterSvc
  {
    public bool AddWords(IEnumerable<KeyValuePair<string, string>> keyValuePairs)
    {
      if (keyValuePairs == null || !keyValuePairs.Any())
        return false;

      return Svc<AutocompleterPlugin>.Plugin.AddWords(keyValuePairs);
    }

    public void Enable()
    {
      Svc<AutocompleterPlugin>.Plugin.Enable();
    }
    public void Disable()
    {
      Svc<AutocompleterPlugin>.Plugin.Disable();
    }
  }
}
