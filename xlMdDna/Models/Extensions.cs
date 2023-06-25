using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace xlMdDna {
    public static class Extensions {
    public static string ThisMonthTxtName => $@"{DateTime.Today.ToString("yyyyMM")}.txt";
    public static string TodayTxtName => $@"{DateTime.Today.ToString("yyyyMMdd")}.txt";
    public static string ThisMonthCsvName => $@"{DateTime.Today.ToString("yyyyMM")}.csv";
    public static string TodayCSVName => $@"{DateTime.Today.ToString("yyyyMMdd")}.csv";
    public static Encoding Enc => Encoding.UTF8;
    public static FileInfo MyExeFi => new FileInfo(Process.GetCurrentProcess().MainModule.FileName);
    public static FileInfo DumpFi => new FileInfo(Path.Combine($"{BaseDi.FullName}/Log", MyExeFi.Name.Replace(".exe", "") + $".{TodayTxtName}"));
    public static FileInfo ErrFi => new FileInfo(Path.Combine($"{BaseDi.FullName}/Exception", MyExeFi.Name.Replace(".exe", "") + $".{TodayTxtName}"));
    public static FileInfo GetAsmFi => new FileInfo(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location));
    //public static System.Configuration.Configuration UserConfig =>
    //    System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal);

    public static DirectoryInfo BaseDi {
      get {
        DirectoryInfo baseDi = new DirectoryInfo(MyExeFi.DirectoryName);
        if (MyExeFi.Name == "dotnet.exe") {
          baseDi = new DirectoryInfo(Directory.GetCurrentDirectory());
        }
        return baseDi;
      }
    }

    public static FileInfo GetFi(string path) {
      if (string.IsNullOrEmpty(path)) {
        return Extensions.DumpFi;
      }
      var fi = new FileInfo(path);
      if (fi.Exists || fi.Directory.Exists) {
        return fi;
      }
      fi = new FileInfo(Path.Combine(BaseDi.FullName, path));
      return fi;
    }

    public static void WriteText(this FileInfo fi, string text, Encoding _enc = null, string _newLineStr = null) {
      if (!fi.Directory.Exists) {
        fi.Directory.Create();
      }

      var newLineStr = (_newLineStr ?? Environment.NewLine);
      text = Regex.Replace(text, "\r\n|\r|\n", newLineStr);

      var enc = (_enc ?? Enc);
      var buf = enc.GetBytes(text);
      using (var fs = new FileStream(fi.FullName, FileMode.Append, FileAccess.Write, FileShare.None, bufferSize: 4096, useAsync: true)) {
        fs.Write(buf, 0, buf.Length);
      };
    }

    public static void DeleteFi(this FileInfo fi) {
      if (fi.Exists) {
        fi.Delete();
      }
    }

    public static Object dumpLock = new object();
    public static T Dump<T>(this T obj, string msg = "", string filePath = null) {
      try {
        var txt = $@"
{DateTime.Now} {msg}
{new string('=', $"{DateTime.Now} {msg}".Length)}
{obj.GetType()}
{new string('-', $"{obj.GetType()}".Length)}
{obj}
{new string('=', $"{DateTime.Now} {msg}".Length)}
";
        FileInfo fi;
        if (!string.IsNullOrEmpty(filePath)) {
          fi = GetFi(filePath);
        } else {
          fi = DumpFi;
        }
        lock (dumpLock) {
          fi.WriteText($"{txt}");
        }
      } catch (Exception) { }
      return obj;
    }

    public static T DebugDump<T>(this T obj, string msg = "", string filePath = null) {
#if DEBUG
      try {
        var txt = $@"
{DateTime.Now} {msg}
{new string('=', $"{DateTime.Now} {msg}".Length)}
{obj.GetType()}
{new string('-', $"{obj.GetType()}".Length)}
{obj}
{new string('=', $"{DateTime.Now} {msg}".Length)}
";
        FileInfo fi;
        if (!string.IsNullOrEmpty(filePath)) {
          fi = GetFi(filePath);
        } else {
          fi = DumpFi;
        }
        lock (dumpLock) {
          Console.WriteLine($"{txt}");
          fi.WriteText($"{txt}");
        }
      } catch (Exception) { }
#endif //DEBUG
      return obj;
    }

   public static string ExceptionStr(this Exception ex)
    {
      return $@"{ex.Message}
            {ex.Source}
            {ex.StackTrace}
            {ex.InnerException?.Message}
            {ex.InnerException?.Source}
            {ex.InnerException?.StackTrace}";
    }

    public static Exception Dump(this Exception ex, string msg = "", string filePath = null) {
      var emsg = $@"{ex.Message}
            {ex.Source}
            {ex.StackTrace}
            {ex.InnerException?.Message}
            {ex.InnerException?.Source}
            {ex.InnerException?.StackTrace}";
      emsg.Dump(msg, filePath: ErrFi.FullName);
      return ex;
    }

    public static T ErrDump<T>(this T obj, string msg = "") {
      return obj.Dump(msg, ErrFi.FullName);
    }

    public static bool TrySetValByName<T>(this T obj, string propName, string val) {
      var noErr = true;
      try {
        var pi = obj.GetType().GetProperty(propName);
        var typ = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;
        if (typ.Name == "TimeSpan") {
          var safeValue = TimeSpan.Parse(val);
          pi.SetValue(obj, safeValue, null);
        } else {
          var safeValue = (val == null) ? null : Convert.ChangeType(val, typ);
          pi.SetValue(obj, Convert.ChangeType(safeValue, typ), null);
        }
      } catch (Exception) {
        noErr = false;
        var msg = $"{propName}:{val}";
        obj.Dump(msg: msg, filePath: ErrFi.FullName);
      }
      return noErr;
    }

    public static string JoinString(this IEnumerable<string> strs) {
      return string.Join(Environment.NewLine, strs);
    }

    public static bool TrySetValByFixedLengthString<T>(this T obj, string str, string propName, int startIndex, int length, int offset = 0) {
      var noErr = true;
      string val = "";
      try {
        val = str.Substring(startIndex, length);
        var pi = obj.GetType().GetProperty(propName);
        //NotSupportNull
        //pi.SetValue(obj, Convert.ChangeType(val, pi.PropertyType), null);
        var typ = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;
        var safeValue = (val == null) ? null : Convert.ChangeType(val, typ);
        pi.SetValue(obj, Convert.ChangeType(safeValue, typ), null);
      } catch (Exception) {
        noErr = false;
        var msg = $"{propName}:{val} [{str}]";
        obj.Dump(msg: msg, filePath: ErrFi.FullName);
      }
      return noErr;
    }

    static public void InfoDialog(string msg = "Infomation", string title = "Infomation") {
      System.Windows.MessageBoxResult result = System.Windows.MessageBox.Show(msg, title, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Exclamation);
    }

    static public bool GetYesNoWithDialog(string msg = "Do you want to continue?", string title = "ConfirmDialog") {
      System.Windows.MessageBoxResult result = System.Windows.MessageBox.Show(msg, title, System.Windows.MessageBoxButton.YesNoCancel, System.Windows.MessageBoxImage.Warning);
      return result == System.Windows.MessageBoxResult.Yes;
    }

    static public string FileOpenDialog(string filter = "全てのファイル(*.*)|*.*", string filename = "", string title = "ファイルを開く") {
      var dialog = new Microsoft.Win32.OpenFileDialog() {
        Title = title,
        Filter = filter,
        FileName = filename,
      };
      if (dialog.ShowDialog() == true) {
        return dialog.FileName;
      }
      return "";
    }

    static public string FileSaveDialog(string filter = "テキストファイル|*.txt", string filename = "", string title = "ファイルを保存") {
      var dialog = new Microsoft.Win32.SaveFileDialog() {
        Title = title,
        Filter = filter,
        FileName = filename,
      };
      if (dialog.ShowDialog() == true) {
        return dialog.FileName;
      }
      return "";
    }

    static public string ReplaceInvalidFileNameChars(string target, string rep = "") {
      //var RgxInvalidFileNameChars =	new Regex( $"[{new String(Path.GetInvalidFileNameChars() )}]" );
      var rgxInvalidFileNameChars = new Regex(@"[<>|:*?/\\]", RegexOptions.Compiled);
      return rgxInvalidFileNameChars.Replace(target, rep);
    }
  }//class
}
