using Reactive.Bindings;
using Svg;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using xlMdDna.Models;
using xlMdDna.Properties;
using MSO = Microsoft.Office.Core;
using MSOXL = Microsoft.Office.Interop.Excel;

namespace xlMdDna.Views {
    /// <summary>
    /// MermaidWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MermaidWindow : Window, IDisposable {
        string mdstr;
        string shapname;
        int emfUse; //1:use
        MSOXL.Application xl;
        MSOXL.Range rng;
        MermaidWindowVM vm;
        Task job;
        
      public MermaidWindow(string mdstr, string shapname, MSOXL.Application xl, MSOXL.Range rng, int emfUse = 1) {
            this.mdstr = mdstr;
            this.shapname = shapname;
            this.emfUse = emfUse;
            this.xl = xl;
            this.rng = rng;
            InitializeComponent();
            this.Show();
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e) {
            try {
                vm = (MermaidWindowVM)this.DataContext;
                vm.checkWindowPosition(this);
                vm.webView = webView;
                webView.CreationProperties = new Microsoft.Web.WebView2.Wpf.CoreWebView2CreationProperties();
                webView.CreationProperties.UserDataFolder = vm.WebViewSaveDirPath.Value;
                await webView.EnsureCoreWebView2Async();
                //webView.NavigateToString($@"<h4> init </h4>");
                vm.WebViewNavigate(mdstr, shapname, xl, rng, emfUse);
            } catch (Exception ex) {
                try { this.Close(); } catch { }
                Extensions.InfoDialog(ex.ExceptionStr(), "Error");
                System.Windows.Forms.Clipboard.SetText(ex.ExceptionStr());
            }
        }

        //private async void WebView_NavigationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs e) {
        //    await webView.ExecuteScriptAsync(MermaidJsString.js.Replace("'", "\""));
        //}

        private void webView_WebMessageReceived(object sender, Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs e) {
            if (e.TryGetWebMessageAsString() == "EnterKeyDown") {
                if (job != null && !job.IsCompleted)
                    job.Dispose();
                job = vm.WebViewPressEnterAsync(this);
            }
        }

        public void Dispose() {
            try { this.Close(); } catch { }
            webView.Dispose();
        }
    }

  public partial class MermaidWindowVM : IDisposable {

    string mdstr;
    string shapname;
    int emfUse;
    MSOXL.Application xl;
    MSOXL.Range rng;
    public ReactiveProperty<string> WebViewSaveDirPath = new ReactiveProperty<string>();
    public Microsoft.Web.WebView2.Wpf.WebView2 webView { get; set; }
    public ReactiveProperty<WindowState> WindState { get; set; } = new ReactiveProperty<WindowState>();
    public ReactiveCommand<System.Windows.Window> CmdWindClosing { get; set; } = new ReactiveCommand<System.Windows.Window>();

    public DirectoryInfo Desktop => new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
    public DirectoryInfo MyDoc => new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
    public DirectoryInfo SaveDir => new DirectoryInfo($@"{MyDoc.FullName}\xlMdDna");
    public FileInfo MermaidJsFi => new FileInfo(System.IO.Path.Combine(SaveDir.FullName, "mermaid.10.2.3.js"));
    public FileInfo MermaidCssFi => new FileInfo(System.IO.Path.Combine(SaveDir.FullName, "mermaid.7.0.11.css"));

    public MermaidWindowVM() {
      setDir();
      subscribe();
    }

    void setDir() {
      WebViewSaveDirPath.Value = SaveDir.FullName;
      putMermaidSrc();
    }

    void subscribe() {
      CmdWindClosing.Subscribe(s => doWindClosing(s));
    }

    void putMermaidSrc() {
      if (!MermaidJsFi.Directory.Exists)
        MermaidJsFi.Directory.Create();
      if (!MermaidJsFi.Exists)
        MermaidJsFi.WriteText(MermaidJsString.js.Replace("'", "\""));
      if (!MermaidCssFi.Exists)
        MermaidCssFi.WriteText(MermaidJsString.css.Replace("'", "\""));
    }

    public async Task WebViewPressEnterAsync(Window wind) {
      var svgStrRaw = await webView.ExecuteScriptAsync("document.getElementById('preview').innerHTML");
      var svgStr = DecodeEncodedNonAsciiCharacters(svgStrRaw);
      svgStr = svgStr.Substring(1, svgStr.Length - 2).Replace("\\\"", "\"");
      //Extensions.InfoDialog(svgStr, "GetSVG");
      var svgPath = $"{SaveDir.FullName}/{shapname}.svg";
      File.WriteAllText(svgPath, textTagRewrite(svgStr), Encoding.UTF8);
      resultPaste(svgPath);
      wind.Close();
    }

    //SVG対応かバージョンで判定(適当...2016のマイナーバージョンのどこかから対応)
    string xlver => xl.Version; //aa.b.cccc.dddd //aa.bまでしかとれない？
    bool IsValidSvg => (emfUse == 0) && double.Parse($"{xlver.Split('.')[0]}.{xlver.Split('.')[1]}") >= 16.0;

    async void resultPaste(string svgPath) {
      var ws = (MSOXL.Worksheet)xl.ActiveSheet;
      try { //前回作成したものがあれば消去
        var tshap = ws.Shapes.Item(shapname);
        tshap.Delete();
      } catch { }

      //エクセルバージョンでSVG対応か判断
      var emf = svg2emf(svgPath);
      var emfPath = emf.Item1;
      var width = emf.Item2;
      var height = emf.Item3;
      MSOXL.Shape shap;
      if (IsValidSvg) {
        shap = ws.Shapes.AddPicture(svgPath, MSO.MsoTriState.msoFalse, MSO.MsoTriState.msoCTrue, 0f, 0f, width, height);
      } else {
        shap = ws.Shapes.AddPicture(emfPath, MSO.MsoTriState.msoFalse, MSO.MsoTriState.msoCTrue, 0f, 0f, width, height);
      }
      shap.Name = shapname;
      shap.Left = float.Parse($"{rng.Offset[0, 1].Left}");
      shap.Top = float.Parse($"{rng.Top}");
    }
    // ActiveSheet.Pictures.Insert(_ "C:\Users\tg30266\Documents\xlMdDna\Book2_Sheet1_A5.emf").Select
    //Selection.ShapeRange.Ungroup.Select

    string textTagRewrite(string buf) {
      var res = buf;

      var rgxFOA = new Regex(
      @"<g transform=[^>]*(-*\d+(?:\.\d+)?), (-*\d+(?:\.\d+)?)\)[^>]* class=""label"">([\s\S]*?)" +
      @"<foreignObject[^>]* height=""(-*\d+(?:\.\d+)?)"" width=""(-*\d+(?:\.\d+)?)""[^>]*>" +
      @"([\s\S]*?)</foreignObject>" +
      @"([\s\S]*?)</g>"
      , RegexOptions.Compiled | RegexOptions.Multiline);
      var matchesFOA = rgxFOA.Matches(buf);
      foreach (Match mFOA in matchesFOA) {
        mFOA.Dump();
        var xs = mFOA.Groups[1];
        var ys = mFOA.Groups[2];
        var x = 0;//double.Parse($"{xs}") * 0.2; //4d;
        var y = 5;//double.Parse($"{ys}") * 1.5; //12d;
        var rgxSN = new Regex(@"<span([\s\S]*?)>([\s\S]*?)</span>", RegexOptions.Compiled | RegexOptions.Multiline);
        var matchesSN = rgxSN.Matches(mFOA.Value);
        foreach (Match mSN in matchesSN) {
          //mSN.Dump();
          var t = mSN.Groups[2];
          var newTag = $@"<text x=""{x}"" y=""{y}"">{t}</text>".Dump();
          //$"===\n{mFO.Value}\n---\n{mSN.Value}".Dump();
          res = res.Replace(mFOA.Value, newTag);
        }
      }

      var rgxFOB = new Regex(@"<foreignObject[^>]* height=""(-*\d+(?:\.\d+)?)"" width=""(-*\d+(?:\.\d+)?)""[^>]*>([\s\S]*?)</foreignObject>", RegexOptions.Compiled | RegexOptions.Multiline);
      var matchesFOB = rgxFOB.Matches(buf);
      foreach (Match mFOB in matchesFOB) {
        //mFO.Dump();
        var xs = mFOB.Groups[1];
        var ys = mFOB.Groups[2];
        var x = double.Parse($"{xs}") * 0.2; //4d;
        var y = double.Parse($"{ys}") * 1.5; //12d;
        var rgxSN = new Regex(@"<span([\s\S]*?)>([\s\S]*?)</span>", RegexOptions.Compiled | RegexOptions.Multiline);
        var matchesSN = rgxSN.Matches(mFOB.Value);
        foreach (Match mSN in matchesSN) {
          //mSN.Dump();
          var t = mSN.Groups[2];
          var newTag = $@"<text x=""{x}"" y=""{y}"">{t}</text>".Dump();
          //	$"===\n{mFO.Value}\n---\n{mSN.Value}".Dump();
          res = res.Replace(mFOB.Value, newTag);
        }
      }
      return res;
    }

    /// <returns>emfPath, imgWidth, imgHeight</returns>
    Tuple<string, int, int> svg2emf(string svgPath) {
      var svgFi = new FileInfo(svgPath);
      var outPathEmf = svgFi.FullName.Replace(svgFi.Extension, ".emf");

      var svg = SvgDocument.Open(svgPath);
      using (var bufg = Graphics.FromHwndInternal(IntPtr.Zero))
      using (var meta = new Metafile(outPathEmf, bufg.GetHdc()))
      using (var g = Graphics.FromImage(meta)) {
        svg.Draw(g);
      }
      var emf = Image.FromFile(outPathEmf);
      //return new Tuple<string, int, int>(outPathEmf, (int)(svg.Bounds.Width*2.55), (int)(svg.Bounds.Height*2.55));
      return new Tuple<string, int, int>(outPathEmf, (int)(emf.Width), (int)(emf.Height));
    }


    string DecodeEncodedNonAsciiCharacters(string value) {
      return Regex.Replace(value, @"\\u(?<Value>[a-zA-Z0-9]{4})",
          m => { return ((char)int.Parse(m.Groups["Value"].Value, System.Globalization.NumberStyles.HexNumber)).ToString(); });
    }

    public void WebViewNavigate(string mdstr, string shapname, MSOXL.Application xl, MSOXL.Range rng, int emfUse = 1) {
      this.mdstr = mdstr;
      this.shapname = shapname;
      this.xl = xl;
      this.rng = rng;
      this.emfUse = emfUse;
      webView.NavigateToString(MMHTML);
      File.WriteAllText(System.IO.Path.Combine(Desktop.FullName, "test.html"), MMHTML);
    }

    private string MMHTML => $@"
<html lang='ja'>
<head>
	<meta charset='utf-8'>
	<!--<meta name='viewport' content='width=device-width, initial-scale=1'>-->
	<title>xlMdDnaPreview</title>
    <!--<script src='https://cdnjs.cloudflare.com/ajax/libs/mermaid/6.0.0/mermaid.min.js'></script>-->
    <!--<link rel='stylesheet' type='text/css' href='https://cdnjs.cloudflare.com/ajax/libs/mermaid/6.0.0/mermaid.min.css'>-->
    <script src='https://cdnjs.cloudflare.com/ajax/libs/mermaid/10.2.3/mermaid.min.js'></script>
    <link rel='stylesheet' type='text/css' href='https://get.cdnpkg.com/mermaid/7.0.10/mermaid.min.css?id=54484'>
    <!--<script src='{MermaidJsFi.FullName}'></script>-->
    <!--<link rel='stylesheet' type='text/css' href='{MermaidCssFi.FullName}'>-->
<style>
  .box {{
    margin: 1em;
    border: 3px solid #F89174;
    border-radius: 15px;
  }}
</style>
</head>
<body>
  <center>

    <div id='preview' class='mermaid'>
{mdstr}
    </div>

    <div class='box'>
      <b>THIRD-PARTY SOFTWARE NOTICES AND INFORMATION</b>
      <a href='https://excel-dna.net/'>https://excel-dna.net/</a><br>
      <a href='https://mermaid.live/'>https://mermaid.live/</a>
    </div>

  </center>

  <script>
window.addEventListener(""keydown"", handleKeydown);
function handleKeydown(event) {{
    var keyCode = event.keyCode;
    // Enter key
    if (keyCode == 13) {{
        window.chrome.webview.postMessage(""EnterKeyDown"");
    }}
}}
  </script>
</body>
</html>
";

        public void checkWindowPosition(System.Windows.Window window) {
            var screen = System.Windows.Forms.Screen.FromHandle(new System.Windows.Interop.WindowInteropHelper(window).Handle);
            var defset = Properties.Settings.Default;
            if (screen.Primary && (
                defset.Left < 0 || defset.Left > screen.Bounds.Width
                || defset.Top < 0 || defset.Top > screen.Bounds.Height
                || defset.Left + defset.Width > screen.Bounds.Width
                || defset.Top + defset.Height > screen.Bounds.Height
                )
            ) {
                defset.Top = 100;
                defset.Left = 100;
                defset.Width = 640;
                defset.Height = 480;
            }
            if (Settings.Default.Maximized) {
                WindState.Value = WindowState.Maximized;
            }
        }

        void doWindClosing(System.Windows.Window window) {
            //保存するときに修正したければ
            //var screen = System.Windows.Forms.Screen.FromHandle(new System.Windows.Interop.WindowInteropHelper(window).Handle);
            saveWindState();
            webView.Dispose();
        }

        void saveWindState() {
            Settings.Default.Maximized = WindState.Value == WindowState.Maximized;
            Settings.Default.Save();
        }

        public void Dispose() {
            xl = null;
        }
    }

}