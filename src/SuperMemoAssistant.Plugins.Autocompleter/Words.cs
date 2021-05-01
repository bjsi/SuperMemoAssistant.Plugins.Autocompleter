using Anotar.Serilog;
using System;
using System.IO;
using System.Linq;

namespace SuperMemoAssistant.Plugins.Autocompleter
{
  public static class Words
  {

    public static readonly char[] PunctuationAndSymbols = new char[]
    {
      '.', '!', '?', ')', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '+', '=', '\\', '/', '<', '>', ','
    };

    private static string FilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"words\top_english_words.txt");
    public static string[] LoadWords()
    {
      try
      {
        string[] lines = File.ReadAllLines(FilePath);
        return lines
          ?.Where(x => x.Length > 3)
          ?.ToArray(); // Only load words longer than 3 letters
      }
      catch (FileNotFoundException)
      {
        LogTo.Error($"Failed to LoadWords because {FilePath} does not exist");
        return null;
      }
      catch (IOException e)
      {
        LogTo.Error($"Exception {e} thrown when attempting to read from {FilePath}");
        return null;
      }
      catch (Exception e)
      {
        LogTo.Error($"Exception {e} thrown when attempting to read from {FilePath}");
        return null;
      }
    }
  }
}
