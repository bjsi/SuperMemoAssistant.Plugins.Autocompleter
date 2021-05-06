using mshtml;
using System.Collections.Generic;
using System.Drawing;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
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

  public static class MeasureWord
  {

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

    public static double GetWordWidth(string word, int fontSize, string fontName)
    {

      if (string.IsNullOrEmpty(word) || fontSize < 0 || string.IsNullOrEmpty(fontName))
        return -1;

      Font stringFont = new Font(fontName, fontSize);
      var stringSize = System.Windows.Forms.TextRenderer.MeasureText(word, stringFont);
      stringFont.Dispose();

      return stringSize == null
        ? -1
        : stringSize.Width;
    }
  }
}
