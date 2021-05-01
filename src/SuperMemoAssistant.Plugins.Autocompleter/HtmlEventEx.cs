using mshtml;
using SuperMemoAssistant.Extensions;
using System;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public interface IControlHtmlEvent
  {
    event EventHandler<IControlHtmlEventArgs> OnEvent;
    void handler(IHTMLEventObj e);
  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  public class HtmlEvent : IControlHtmlEvent
  {
    public event EventHandler<IControlHtmlEventArgs> OnEvent;

    [DispId(0)]
    public void handler(IHTMLEventObj e)
    {
      if (OnEvent != null)
        OnEvent(this, new IControlHtmlEventArgs(e));
    }
  }

  public class IControlHtmlEventArgs
  {
    public IHTMLEventObj EventObj { get; set; }
    public IControlHtmlEventArgs(IHTMLEventObj EventObj)
    {
      this.EventObj = EventObj;
    }
  }

  public static class HtmlEventEx
  {
    public enum EventType
    {

      onkeydown,
      onclick,
      ondblclick,
      onkeypress,
      onkeyup,
      onmousedown,
      onmousemove,
      onmouseout,
      onmouseover,
      onmouseup,
      onselectstart,
      onbeforecopy,
      onbeforecut,
      onbeforepaste,
      oncontextmenu,
      oncopy,
      oncut,
      ondrag,
      ondragstart,
      ondragend,
      ondragenter,
      ondragleave,
      ondragover,
      ondrop,
      onfocus,
      onlosecapture,
      onpaste,
      onpropertychange,
      onreadystatechange,
      onresize,
      onactivate,
      onbeforedeactivate,
      oncontrolselect,
      ondeactivate,
      onmouseenter,
      onmouseleave,
      onmove,
      onmoveend,
      onmovestart,
      onpage,
      onresizeend,
      onresizestart,
      onfocusin,
      onfocusout,
      onmousewheel,
      onbeforeeditfocus,
      onafterupdate,
      onbeforeupdate,
      ondataavailable,
      ondatasetchanged,
      ondatasetcomplete,
      onerrorupdate,
      onfilterchange,
      onhelp,
      onrowenter,
      onrowexit,
      onlayoutcomplete,
      onblur,
      onrowsdelete,
      onrowsinserted,

    }

    public static bool SubscribeTo(this IHTMLElement2 element, EventType eventType, IControlHtmlEvent handlerObj)
    {
      try
      {

        return element == null
          ? false
          : element.attachEvent(eventType.Name(), handlerObj);

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return false;
    }

    public static void UnsubscribeFrom(this IHTMLElement2 element, EventType eventType, IControlHtmlEvent handlerObj)
    {
      try
      {
        element?.detachEvent(eventType.Name(), handlerObj);
      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }
    }

    public static bool SubscribeTo(this IHTMLElement2 element, string @event, object pdisp)
    {

      try
      {

        return element == null
          ? false
          : element.attachEvent(@event, pdisp);

      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }

      return false;
    }

    public static void UnsubscribeFrom(this IHTMLElement2 element, string @event, object pdisp)
    {
      try
      {
        element?.detachEvent(@event, pdisp);
      }
      catch (RemotingException) { }
      catch (UnauthorizedAccessException) { }
    }
  }
}
