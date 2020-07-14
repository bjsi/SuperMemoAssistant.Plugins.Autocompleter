using Anotar.Serilog;
using mshtml;
using SuperMemoAssistant.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace SuperMemoAssistant.Plugins.Autocompleter
{

  public enum HtmlCommand
  {
    FontName,
    FontSize,
  }

  public static class SelObjEx
  {

    public static int QueryFontSize(this IHTMLTxtRange selObj)
    {
      return selObj.IsNull()
        ? -1
        : (int)selObj.QueryValueSelObj(HtmlCommand.FontSize);
    }

    public static string QueryFontName(this IHTMLTxtRange selObj)
    {
      return selObj.IsNull()
        ? null
        : (string) selObj.QueryValueSelObj(HtmlCommand.FontName);
    }


    [LogToErrorOnException]
    private static object QueryValueSelObj(this IHTMLTxtRange selObj, HtmlCommand command)
    {
      try
      {

        if (selObj.IsNull())
          return null;

        // ensure command is valid and enabled
        if (!selObj.queryCommandSupported(command.Name()))
          return null;

        if (!selObj.queryCommandEnabled(command.Name()))
          return null;

        return selObj.queryCommandValue(command.Name());

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return null;

    }

  }
}
