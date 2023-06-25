using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using xlMdDna.Views;
using MSOXL = Microsoft.Office.Interop.Excel;

namespace xlMdDna.Models
{

  public class ExcelFunctionEntry : IDisposable
  {

    static MermaidWindow mmWind;

    [ExcelFunction(Name = "Mermaid", Description = "About xlMdDna")]
    public static string Mermaid(dynamic[,] args, int emfUse = 0)
    { //※optionalのデフォ値は効かない。//エクセル関数の引数小着磁は型のデフォ値が渡される（たぶん）
      try
      {
        var xl = (MSOXL.Application)ExcelDnaUtil.Application;
        var caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
        var wb = (MSOXL.Workbook)xl.ActiveWorkbook;
        var ws = (MSOXL.Worksheet)xl.ActiveSheet;
        var rng = (MSOXL.Range)ws.Cells[caller.RowFirst + 1, caller.ColumnFirst + 1];
        var shapName = $"{wb.Name}_{ws.Name}_{rng.Address[false, false]}";
        var buf = getArgsString(args);
        var mdstr = string.Join("\n", buf).Replace("\u00A0", " ");
        if (mmWind != null)
          mmWind.Dispose();
        mmWind = new MermaidWindow(mdstr, shapName, xl, rng, emfUse);
      } catch (Exception ex)
      {
        System.Windows.Forms.Clipboard.SetText($"Err: mermaidFail\n{ex.Message}");
        return "NG";
      }

      return "OK";
    }

    private static IEnumerable<string> getArgsString(object[,] args)
    {
      var yLen = args.GetLength(0);
      var xLen = args.GetLength(1);
      var line = "";
      var str = "";
      var rgx = new Regex(@"^(\(|\[|\{)");
      for (var y = 0; y < yLen; y++)
      {
        line = "";
        for (var x = 0; x < xLen; x++)
        {
          try
          {
            if ((str = args[y, x].ToString()) != "ExcelDna.Integration.ExcelEmpty")
              line += (rgx.IsMatch(str) ? "" : " ") + str;
          } catch (Exception ex)
          {
            System.Windows.Forms.Clipboard.SetText($"Err: ReadCellFail\n{ex.Message}\n{args[y, x]}");
          }
        }
        yield return line;
      }
    }

    private static string sjisToUtf(string sjisStr)
    {
      Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
      byte[] bytesData = System.Text.Encoding.UTF8.GetBytes(sjisStr);
      Encoding utf8Enc = Encoding.GetEncoding("UTF-8");
      return utf8Enc.GetString(bytesData);
    }

    public void Dispose()
    {
      if (mmWind != null)
        mmWind.Dispose();
    }
  }
}
