using Forge.Forms.Annotations;
using Newtonsoft.Json;
using SuperMemoAssistant.Services.UI.Configuration;
using SuperMemoAssistant.Sys.ComponentModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  [Form(Mode = DefaultFields.None)]
  [Title("Dictionary Settings",
       IsVisible = "{Env DialogHostContext}")]
  [DialogAction("cancel",
      "Cancel",
      IsCancel = true)]
  [DialogAction("save",
      "Save",
      IsDefault = true,
      Validates = true)]
  public class AutocompleterCfg : CfgBase<AutocompleterCfg>, INotifyPropertyChangedEx
  {

    [Title("Autocompleter Plugin")]
    [Heading("By Jamesb | Experimental Learning")]

    [Heading("Features:")]
    [Text(@"- Adds IDE-style autocompletion popups to SuperMemo HTML components.")]

    [Heading("General Settings")]

    [Field(Name = "Minimum search result word length?")]
    public int MinSearchLength { get; set; } = 3;

    [Field(Name = "Max Results?")]
    public int MaxResults { get; set; } = 10;

    [JsonIgnore]
    public bool IsChanged { get; set; }

    public override string ToString()
    {
      return "Autocompleter Settings";
    }

    public event PropertyChangedEventHandler PropertyChanged;

  }
}
