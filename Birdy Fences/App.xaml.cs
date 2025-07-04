using Birdy_Browser;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shell;
using System.Windows.Threading;

namespace Birdy_Fences
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpWindowClass, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        const int GWL_HWNDPARENT = -8;

        IntPtr hprog = FindWindowEx(
            FindWindowEx(
                FindWindow("Progman", "Program Manager"),
                IntPtr.Zero, "SHELLDLL_DefView", ""
            ),
            IntPtr.Zero, "SysListView32", "FolderView"
        );

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            string userdir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            Directory.CreateDirectory(userdir + "\\Birdy Fences");
            // fences.json exist but contains only []
            if (!File.Exists(userdir + "\\Birdy Fences\\fences.json") || File.ReadAllText(userdir + "\\Birdy Fences\\fences.json").Trim() == "[]")
            {
                File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", "[{\"Title\":\"Welcome to BirdyFences\",\"X\":500,\"Y\":500,\"Width\":500,\"Height\":200,\"ItemsType\":\"Data\",\"isLocked\":false,\"Items\":[{\"Filename\":\"" + userdir.Replace("\\", "\\\\") + "\\\\Birdy Fences\\\\WELCOME.txt\"}]}]");
                File.WriteAllText(userdir + "\\Birdy Fences\\WELCOME.txt", @"Welcome to Birdy Fences!

This is a simple application that allows you to create fences on your desktop to organize your files and folders.
The Fences are draggable and sizable to help you organize. Drag and drop files and folders into the fences to add them.

To create a new fence, right click on the Title of the fence, then New Fence
To remove a fence, right click again in the Title of the fence, then Remove Fence
To create a Portal Fence, right click again in the Title of the fence then New Portal Fence, and select a folder to import all shortcuts
To Lock/Unlock a fence, right click on the Title of the fence, then Lock Fence
To edit title of the fence, double click on the title, then type the new title and press Enter

Terminologies:
Portal Fence - are files and shortcuts that exists on the selected Portal Folder");
            }
            dynamic? fencedata = Newtonsoft.Json.JsonConvert.DeserializeObject(File.ReadAllText(userdir + "\\Birdy Fences\\fences.json"));
            void createFence(dynamic fence) {
                bool isLocked = fence["isLocked"] != null && (bool)fence["isLocked"];
                DockPanel dp = new();
                Border cborder = new() { Background = new SolidColorBrush(Color.FromArgb(100, 0, 0, 0)), CornerRadius = new CornerRadius(6), Child = dp };
                ContextMenu cm = new();
                MenuItem miNF = new() { Header = "New Fence" };
                cm.Items.Add(miNF);
                MenuItem miNP = new() { Header = "New Portal Fence" };
                cm.Items.Add(miNP);
                MenuItem miRF = new() { Header = "Remove Fence" };
                cm.Items.Add(miRF);
                cm.Items.Add(new Separator());
                MenuItem miLF = new() { Header = "Lock Fence", IsCheckable = true, IsChecked = isLocked };
                cm.Items.Add(miLF);

                Window win = new() { ContextMenu = cm, AllowDrop = true, AllowsTransparency = true, Background = Brushes.Transparent, Title = fence["Title"], ShowInTaskbar = false, WindowStyle = WindowStyle.None, Content = cborder, ResizeMode = isLocked ? ResizeMode.NoResize : ResizeMode.CanResize, Width = fence["Width"], Height = fence["Height"], Top = fence["Y"], Left = fence["X"] };
                HideAltTab.VirtualDesktopHelper.MakeWindowPersistent(win);
                miLF.Click += (sender, e) =>
                {
                    // Toggle fence lock: disables/enables resizing the fence
                    isLocked = !isLocked;
                    win.ResizeMode = isLocked ? ResizeMode.NoResize : ResizeMode.CanResize;
                    fence["isLocked"] = isLocked;

                    File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                };
                miRF.Click += (sender, e) => {
                    fence.Remove();
                    win.Close();
                    File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                };
                miRF.Click += (sender, e) => {
                    if (!isLocked)
                    {
                        fence.Remove();
                        win.Close();
                        File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                        cm.Items.Refresh();
                    }
                };
                miNF.Click += (sender, e) => {
                    Newtonsoft.Json.Linq.JObject fnc = new(new Newtonsoft.Json.Linq.JProperty("Title", "New Fence"), new Newtonsoft.Json.Linq.JProperty("Width", 300), new Newtonsoft.Json.Linq.JProperty("Height", 150), new Newtonsoft.Json.Linq.JProperty("X", 0), new Newtonsoft.Json.Linq.JProperty("Y", 0), new Newtonsoft.Json.Linq.JProperty("ItemsType", "Data"), new Newtonsoft.Json.Linq.JProperty("Items", new Newtonsoft.Json.Linq.JArray()));
                    fencedata.Add(fnc);
                    createFence(fnc);
                    File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                };
                miNP.Click += (sender, e) => {
                    using var dialog = new System.Windows.Forms.FolderBrowserDialog
                    {
                        Description = "Select Folder For Portal",
                        UseDescriptionForTitle = true,
                        ShowNewFolderButton = true
                    };
                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        Newtonsoft.Json.Linq.JObject fnc = new(new Newtonsoft.Json.Linq.JProperty("Title", "New Fence"), new Newtonsoft.Json.Linq.JProperty("Width", 300), new Newtonsoft.Json.Linq.JProperty("Height", 150), new Newtonsoft.Json.Linq.JProperty("X", 0), new Newtonsoft.Json.Linq.JProperty("Y", 0), new Newtonsoft.Json.Linq.JProperty("ItemsType", "Portal"), new Newtonsoft.Json.Linq.JProperty("Items", dialog.SelectedPath));

                        fencedata.Add(fnc);
                        createFence(fnc);
                        File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                    }
                };
                WindowChrome.SetWindowChrome(win, new WindowChrome() { CaptionHeight = 0, ResizeBorderThickness = new Thickness(5) });
                Label titlelabel = new() { Content = (string)fence["Title"], Background = new SolidColorBrush(Color.FromArgb(20, 0, 0, 0)), Foreground = Brushes.White, HorizontalContentAlignment = HorizontalAlignment.Center };
                dp.Children.Add(titlelabel);
                TextBox titletb = new() { HorizontalContentAlignment = HorizontalAlignment.Center, Visibility = Visibility.Collapsed };
                dp.Children.Add(titletb);
                titlelabel.MouseDown += (object sender, MouseButtonEventArgs e) => {
                    if (e.ClickCount == 1)
                    {
                        if (e.LeftButton == MouseButtonState.Pressed && !isLocked)
                        {
                            win.DragMove();
                        }
                    }
                    else
                    {
                        if (isLocked) return;
                        titlelabel.Visibility = Visibility.Collapsed;
                        titletb.Visibility = Visibility.Visible;
                        titletb.Text = (string)titlelabel.Content;
                    }
                };
                titletb.KeyDown += (object sender, KeyEventArgs e) => {
                    if (e.Key == Key.Enter)
                    {
                        titlelabel.Visibility = Visibility.Visible;
                        titletb.Visibility = Visibility.Collapsed;
                        titlelabel.Content = titletb.Text;
                        fence["Title"] = titletb.Text;
                        File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));

                    }
                    else if (e.Key == Key.Escape)
                    {
                        titlelabel.Visibility = Visibility.Visible;
                        titletb.Visibility = Visibility.Collapsed;
                    }
                };
                titlelabel.MouseUp += (object sender, MouseButtonEventArgs e) => {
                    fence["Y"] = win.Top;
                    fence["X"] = win.Left;
                    File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                };
                win.SizeChanged += (sender, e) => {
                    fence["Width"] = win.ActualWidth;
                    fence["Height"] = win.ActualHeight;
                    fence["Y"] = win.Top;
                    fence["X"] = win.Left;
                    File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                };
                DockPanel.SetDock(titlelabel, Dock.Top);
                DockPanel.SetDock(titletb, Dock.Top);
                WrapPanel wpcont = new() { AllowDrop = true };
                void addicon(dynamic icon)
                {
                    StackPanel sp = new() { Margin = new Thickness(5) };
                    sp.Width = 60;
                    ContextMenu mn = new();
                    MenuItem miE = new() { Header = "Edit" };
                    MenuItem miM = new() { Header = "Move.." };
                    MenuItem miRemove = new() { Header = "Remove" };
                    miRemove.Click += (sender, e) =>
                    {
                        if (isLocked) return;
                        icon.Remove();
                        wpcont.Children.Remove(sp);
                        File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                    };
                    mn.Items.Add(miE);
                    mn.Items.Add(miM);
                    mn.Items.Add(miRemove);
                    sp.ContextMenu = mn;
                    Image ico = new() { Width = 40, Height = 40, Margin = new Thickness(5) };
                    try
                    {
                        if (icon["DisplayIcon"] == null)
                        {
                            var extractedIcon = System.Drawing.Icon.ExtractAssociatedIcon((string)icon["Filename"]);
                            if (extractedIcon != null)
                            {
                                ico.Source = extractedIcon.ToImageSource();
                            }
                        }
                        else
                        {
                            ico.Source = new BitmapImage(new Uri((string)icon["DisplayIcon"], UriKind.Relative));
                        }
                    }
                    catch
                    { }
                    sp.Children.Add(ico);
                    TextBlock lbl = new() { TextWrapping = TextWrapping.Wrap, TextTrimming = TextTrimming.CharacterEllipsis, HorizontalAlignment = HorizontalAlignment.Center, Foreground = Brushes.White };
                    lbl.MaxHeight = (lbl.FontSize * 1.5) + (lbl.Margin.Top * 2);
                    if (icon["DisplayName"] == null)
                    {
                        lbl.Text = new FileInfo((string)icon["Filename"]).Name;
                    }
                    else
                    {
                        lbl.Text = (string)icon["DisplayName"];
                    }
                    sp.Children.Add(lbl);
                    miM.Click += (sender, e) => {
                        if (isLocked) return;
                        StackPanel cnt = new();
                        Window wwin = new() { Title = "Move " + (string)icon["Filename"], Content = cnt, Width = 300, Height = 100, WindowStartupLocation = WindowStartupLocation.CenterScreen };
                        ComboBox lv = new();
                        foreach (dynamic icn in fence["Items"])
                        {
                            //StackPanel cc = new() { Orientation = Orientation.Horizontal};
                            //cc.Children.Add(new Image() { Source = ico.Source });
                            //cc.Children.Add(new Label() { Content = lbl.Text });
                            lv.Items.Add(icn["Filename"]);
                        }
                        cnt.Children.Add(lv);
                        Button btn = new() { Content = "Move" };
                        cnt.Children.Add(btn);
                        btn.Click += (sender, e) => {
                            int id = wpcont.Children.IndexOf(sp);
                            dynamic olddata = fence["Items"][lv.SelectedIndex];
                            fence["Items"][lv.SelectedIndex] = fence["Items"][id];
                            fence["Items"][id] = olddata;
                            File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                            initcontent();
                            wwin.Close();
                        };
                        wwin.ShowDialog();
                    };
                    miE.Click += (sender, e) => {
                        if (isLocked) return;
                        StackPanel cnt = new();
                        Window wwin = new() { Title = "Edit " + (string)icon["Filename"], Content = cnt, Width = 450, Height = 200, WindowStartupLocation = WindowStartupLocation.CenterScreen };
                        TextBox createsec(string name, string defaulval)
                        {
                            DockPanel dpp = new();
                            Label lbl = new() { Content = name };
                            dpp.Children.Add(lbl);
                            TextBox tbb = new() { Text = defaulval };
                            dpp.Children.Add(tbb);
                            cnt.Children.Add(dpp);
                            return tbb;
                        };
                        int id = wpcont.Children.IndexOf(sp);
                        TextBox tbDN = createsec("Display Name", fence["Items"][id]["DisplayName"] == null ? "{AUTONAME}" : fence["Items"][id]["DisplayName"]);
                        TextBox tbDI = createsec("Display Icon", fence["Items"][id]["DisplayIcon"] == null ? "{AUTOICON}" : fence["Items"][id]["DisplayIcon"]);
                        Button btn = new() { Content = "Apply" };
                        btn.Click += (sender, e) => {
                            if (tbDN.Text == "{AUTONAME}")
                            {
                                try
                                {
                                    fence["Items"][id]["DisplayName"].Remove();
                                }
                                catch { }
                            }
                            else
                            {
                                fence["Items"][id]["DisplayName"] = tbDN.Text;
                                lbl.Text = tbDN.Text;
                            }
                            if (tbDI.Text == "{AUTOICON}")
                            {
                                try
                                {
                                    fence["Items"][id]["DisplayIcon"].Remove();
                                }
                                catch { }
                            }
                            else
                            {
                                fence["Items"][id]["DisplayIcon"] = tbDI.Text;
                                ico.Source = new BitmapImage(new Uri(tbDN.Text));
                            }
                            File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                            wwin.Close();
                        };
                        cnt.Children.Add(btn);
                        wwin.ShowDialog();
                    };
                    var p = new Process();
                    p.StartInfo = new ProcessStartInfo((string)icon["Filename"])
                    {
                        UseShellExecute = true
                    };
                    new ClickEventAdder(sp).Click += (sender, e) => {
                        p.Start();
                    };
                    wpcont.Children.Add(sp);
                };
                win.DragOver += (object sender, DragEventArgs e) => {
                    e.Effects = DragDropEffects.Copy | DragDropEffects.Move;
                    //e.Handled = true;
                };
                win.DragEnter += (object sender, DragEventArgs e) => {
                    e.Effects = DragDropEffects.Copy | DragDropEffects.Move;
                    //e.Handled = true;
                };
                win.Drop += (object sender, DragEventArgs e) => {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    foreach (string dt in files)
                    {
                        Newtonsoft.Json.Linq.JObject icon = new(new Newtonsoft.Json.Linq.JProperty("Filename", dt));
                        fence["Items"].Add(icon);
                        addicon(icon);
                    }
                    File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
                };
                void initcontent()
                {
                    wpcont.Children.Clear();
                    if (fence["ItemsType"] == "Data")
                    {
                        foreach (dynamic icon in fence["Items"])
                        {
                            addicon(icon);
                        }
                    }else if (fence["ItemsType"] == "Portal")
                    {
                        string dpath = (string)fence["Items"];
                        string[] dirs = Directory.GetDirectories(dpath);
                        foreach (string dir in dirs)
                        {
                            Newtonsoft.Json.Linq.JObject icon = new(new Newtonsoft.Json.Linq.JProperty("Filename", dir), new Newtonsoft.Json.Linq.JProperty("DisplayIcon", "folder-White.png"));
                            addicon(icon);
                        }
                        string[] files = Directory.GetFiles(dpath);
                        foreach (string file in files)
                        {
                            Newtonsoft.Json.Linq.JObject icon = new(new Newtonsoft.Json.Linq.JProperty("Filename", file));
                            addicon(icon);
                        }
                    }
                }
                
                initcontent();
                ScrollViewer wpcontscr = new() { Content = wpcont, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
                dp.Children.Add(wpcontscr);
                win.Show();
                win.Loaded += (sender,e) => SetWindowLong(new WindowInteropHelper(win).Handle, GWL_HWNDPARENT, hprog);
            }
            if (fencedata != null)
            {
                foreach (dynamic fence in fencedata)
                {
                    createFence(fence);
                }
            }
        }
    }

    internal static class IconUtilities
    {
        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteObject(IntPtr hObject);

        public static ImageSource ToImageSource(this System.Drawing.Icon icon)
        {
            System.Drawing.Bitmap bitmap = icon.ToBitmap();
            IntPtr hBitmap = bitmap.GetHbitmap();

            ImageSource wpfBitmap = Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());

            if (!DeleteObject(hBitmap))
            {
                throw new Win32Exception();
            }

            return wpfBitmap;
        }
    }

    internal static class HideAltTab
    {
        public static class VirtualDesktopHelper
        {
            [DllImport("user32.dll")]
            static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

            [DllImport("user32.dll")]
            static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

            [DllImport("shell32.dll")]
            static extern int SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);

            [DllImport("user32.dll", SetLastError = true)]
            static extern IntPtr FindWindowEx(IntPtr hP, IntPtr hC, string sC, string? sW);

            const int GWL_EXSTYLE = -20;
            const int WS_EX_TOOLWINDOW = 0x00000080;

            public static void MakeWindowPersistent(Window window)
            {
                window.SourceInitialized += (s, e) =>
                {
                    var helper = new WindowInteropHelper(window);
                    helper.EnsureHandle();
                    IntPtr hWnd = helper.Handle;

                    // Set window owner to SHELLDLL_DefView for hiding in Alt+Tab and making it persistent
                    IntPtr progman = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "Progman", null);
                    IntPtr shellView = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);
                    helper.Owner = shellView;

                    // Force a fake AppUserModelID (optional, for virtual desktop persistence)
                    SetCurrentProcessExplicitAppUserModelID("Birdy.Fence");

                    // Set TOOLWINDOW style to hide from Alt+Tab
                    var exStyle = (int)GetWindowLongPtr(hWnd, GWL_EXSTYLE);
                    exStyle |= WS_EX_TOOLWINDOW;
                    SetWindowLongPtr(hWnd, GWL_EXSTYLE, (IntPtr)exStyle);
                };
            }
        }
    }
}
