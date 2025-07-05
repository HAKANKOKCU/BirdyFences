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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify;

namespace Birdy_Fences
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static string userdir = "";
        public static List<Fence> fencedata = new();   

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            userdir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            Directory.CreateDirectory(userdir + "\\Birdy Fences");
            Directory.CreateDirectory(userdir + "\\Birdy Fences\\Shortcuts");
            // fences.json exist but contains only []
            if (!File.Exists(userdir + "\\Birdy Fences\\fences.json") || File.ReadAllText(userdir + "\\Birdy Fences\\fences.json").Trim() == "[]")
            {
                File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", "[{\"Title\":\"Welcome to BirdyFences\",\"X\":500,\"Y\":500,\"Width\":500,\"Height\":200,\"ItemsType\":\"Data\",\"isLocked\":false,\"Items\":[{\"Filename\":\"" + userdir.Replace("\\", "\\\\") + "\\\\Birdy Fences\\\\WELCOME.txt\"}]}]");
                File.WriteAllText(userdir + "\\Birdy Fences\\WELCOME.txt", @"Welcome to Birdy Fences!

This is a simple application that allows you to create fences on your desktop to organize your files and folders.
The Fences are draggable and sizable to help you organize. Drag and drop files and folders into the fences to add them.

To create a new fence, right click on the Title of a fence, then New Fence
To remove a fence, right click in the Title of the fence, then Remove Fence
To create a Portal Fence, right click in the Title of the fence then New Portal Fence, and select a folder to show it in the fence.
To Lock/Unlock a fence, right click on the Title of the fence, then Lock Fence
To edit title of the fence, double click on the title, then type the new title and press Enter

Terminologies:
Portal Fence - are files and shortcuts that exists on the selected portal folder");
            }
            fencedata = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Fence>>(File.ReadAllText(userdir + "\\Birdy Fences\\fences.json")) ?? new();
            if (fencedata == null) fencedata = new();
            
            if (fencedata != null)
            {
                foreach (Fence fence in fencedata)
                {
                    fence.init();
                }
            }
        }

        public static void save()
        {
            File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", Newtonsoft.Json.JsonConvert.SerializeObject(fencedata));
        }
    }

    public class Fence
    {
        const int GWL_HWNDPARENT = -8;
        [DllImport("user32.dll", SetLastError = true)]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpWindowClass, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        IntPtr hprog = FindWindowEx(
            FindWindowEx(
                FindWindow("Progman", "Program Manager"),
                IntPtr.Zero, "SHELLDLL_DefView", ""
            ),
            IntPtr.Zero, "SysListView32", "FolderView"
        );

        public string Title = "New fence";
        public bool isLocked = false;
        public double X = 500;
        public double Y = 500;
        public double Width = 500;
        public double Height = 300;
        public string ItemsType = "Data";
        public object Items = new List<FenceItem>();
        public byte[] fencecolor = [100, 0, 0, 0]; //ARGB
        DockPanel dp = new();
        Border cborder = new() { CornerRadius = new CornerRadius(6) };
        ContextMenu cm = new();
        MenuItem miNF = new() { Header = "New Fence" };
        MenuItem miNP = new() { Header = "New Portal Fence" };
        MenuItem miRF = new() { Header = "Remove Fence" };
        MenuItem miC = new() { Header = "Color" };
        MenuItem miLF = new() { Header = "Lock Fence", IsCheckable = true };
        Window win = new() { AllowDrop = true, AllowsTransparency = true, Background = Brushes.Transparent, ShowInTaskbar = false, WindowStyle = WindowStyle.None };
        Label titlelabel = new() { Background = new SolidColorBrush(Color.FromArgb(20, 0, 0, 0)), Foreground = Brushes.White, HorizontalContentAlignment = HorizontalAlignment.Center };
        WrapPanel wpcont = new() { AllowDrop = true };

        public void init()
        {
            cm.Items.Add(miNF);
            cm.Items.Add(miNP);
            cm.Items.Add(miRF);
            cm.Items.Add(new Separator());
            cm.Items.Add(miC);
            cm.Items.Add(miLF);
            HideAltTab.VirtualDesktopHelper.MakeWindowPersistent(win);
            win.Content = cborder;
            win.Width = Width;
            win.Height = Height;
            win.Top = Y;
            win.Left = X;
            win.ContextMenu = cm;
            cborder.Child = dp;

            cm.Opened += (e,s) => {
                miLF.IsChecked = isLocked;
                miRF.IsEnabled = !isLocked;
            };

            miLF.Click += (sender, e) =>
            {
                // Toggle fence lock: disables/enables resizing the fence
                isLocked = !isLocked;
                applyFenceSettings();
                App.save();
            };
            miC.Click += (s, e) => {
                // Color picker
                StackPanel mcont = new();
                Window win = new() { Title = "Choose color", ResizeMode = ResizeMode.NoResize, Width = 300, Height = 200, WindowStartupLocation = WindowStartupLocation.CenterScreen, Content = mcont };
                Slider addslider(string name)
                {
                    DockPanel cont = new();
                    Label lbl = new() { Content = name };
                    Slider sldr = new() { Maximum = 255 };
                    cont.Children.Add(lbl);
                    cont.Children.Add(sldr);
                    mcont.Children.Add(cont);
                    return sldr;
                }
                
                Slider redslider = addslider("Red");
                Slider greenslider = addslider("Green");
                Slider blueslider = addslider("Blue");
                Slider alphaslider = addslider("Alpha");
                redslider.Value = fencecolor[1];
                greenslider.Value = fencecolor[2];
                blueslider.Value = fencecolor[3];
                alphaslider.Value = fencecolor[0];

                redslider.ValueChanged += (s, e) => {
                    fencecolor[1] = (byte)redslider.Value;
                    applyFenceSettings();
                    App.save();
                };

                greenslider.ValueChanged += (s, e) => {
                    fencecolor[2] = (byte)greenslider.Value;
                    applyFenceSettings();
                    App.save();
                };

                blueslider.ValueChanged += (s, e) => {
                    fencecolor[3] = (byte)blueslider.Value;
                    applyFenceSettings();
                    App.save();
                };

                alphaslider.ValueChanged += (s, e) => {
                    fencecolor[0] = (byte)alphaslider.Value;
                    applyFenceSettings();
                    App.save();
                };

                win.ShowDialog();
            };
            miRF.Click += (sender, e) => {
                // Remove fence
                if (!isLocked)
                {
                    App.fencedata.Remove(this);
                    win.Close();
                    App.save();
                    cm.Items.Refresh();
                }
            };
            miNF.Click += (sender, e) => {
                var fnc = new Fence();
                App.fencedata.Add(fnc);
                fnc.init();
                App.save();
            };
            miNP.Click += (sender, e) => {
                // New portal fence
                using var dialog = new System.Windows.Forms.FolderBrowserDialog
                {
                    Description = "Select folder for portal",
                    UseDescriptionForTitle = true,
                    ShowNewFolderButton = true
                };
                
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var fnc = new Fence()
                    {
                        ItemsType = "Portal",
                        Items = dialog.SelectedPath
                    };
                    App.fencedata.Add(fnc);
                    fnc.init();
                    App.save();
                }
            };
            WindowChrome.SetWindowChrome(win, new WindowChrome() { CaptionHeight = 0, ResizeBorderThickness = new Thickness(5) });

            titlelabel.Content = Title;
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
                    titletb.Focus();
                    titletb.SelectAll();
                }
            };
            titletb.KeyDown += (object sender, KeyEventArgs e) => {
                if (e.Key == Key.Enter)
                {
                    titlelabel.Visibility = Visibility.Visible;
                    titletb.Visibility = Visibility.Collapsed;
                    Title = titletb.Text;
                    applyFenceSettings();
                    App.save();

                }
                else if (e.Key == Key.Escape)
                {
                    titlelabel.Visibility = Visibility.Visible;
                    titletb.Visibility = Visibility.Collapsed;
                }
            };
            titlelabel.MouseUp += (object sender, MouseButtonEventArgs e) => {
                Y = win.Top;
                X = win.Left;
                App.save();
            };
            win.SizeChanged += (sender, e) => {
                Width = win.ActualWidth;
                Height = win.ActualHeight;
                Y = win.Top;
                X = win.Left;
                App.save();
            };
            DockPanel.SetDock(titlelabel, Dock.Top);
            DockPanel.SetDock(titletb, Dock.Top);
            
            win.DragOver += (object sender, DragEventArgs e) => {
                e.Effects = DragDropEffects.Copy | DragDropEffects.Move;
                //e.Handled = true;
            };
            win.DragEnter += (object sender, DragEventArgs e) => {
                e.Effects = DragDropEffects.Copy | DragDropEffects.Move;
                //e.Handled = true;
            };
            win.Drop += (object sender, DragEventArgs e) => {
                var items = geticons();
                if (isLocked || items == null)
                {
                    return;
                }

                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string dt in files)
                {
                    string extension = Path.GetExtension(dt).ToLower();
                    string filepath = dt;
                    string displayname = "{AUTONAME}";
                    if (extension == ".lnk" || extension == ".url")
                    {
                        string filename = Path.GetFileName(dt);
                        displayname = Path.GetFileNameWithoutExtension(dt);
                        int c = 0;
                        do
                        {
                            filename = c + Path.GetFileName(filepath);
                            filepath = App.userdir + "\\Birdy Fences\\Shortcuts\\" + filename;
                            ++c;
                        } while (File.Exists(filepath));
                        File.Copy(dt, filepath);
                    }
                    var icon = new FenceItem() { Filename = filepath, DisplayName = displayname };
                    items.Add(icon);
                    addicon(icon);
                }
                App.save();
            };


            initcontent();
            ScrollViewer wpcontscr = new() { Content = wpcont, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            dp.Children.Add(wpcontscr);
            win.Show();
            win.Loaded += (sender, e) => SetWindowLong(new WindowInteropHelper(win).Handle, GWL_HWNDPARENT, hprog);

            applyFenceSettings();
        }

        public void initcontent()
        {
            wpcont.Children.Clear();
            if (ItemsType == "Data")
            {
                var items = geticons();
                if (items == null)
                {
                    return;
                }

                foreach (FenceItem icon in items)
                {
                    addicon(icon);
                }
            }
            else if (ItemsType == "Portal")
            {
                string dpath = (string)Items;
                string[] dirs = Directory.GetDirectories(dpath);
                foreach (string dir in dirs)
                {
                    FenceItem icon = new FenceItem() { Filename = dir };
                    addicon(icon);
                }
                string[] files = Directory.GetFiles(dpath);
                foreach (string file in files)
                {
                    if (Path.GetFileName(file).ToLower() == "desktop.ini") continue;
                    FenceItem icon = new FenceItem() { Filename = file };
                    addicon(icon);
                }
            }
        }

        void addicon(FenceItem icon)
        {
            Button btn = new() { Margin = new Thickness(5), Style = (Style)Application.Current.Resources["asbsLight"], Background = Brushes.Transparent, BorderThickness = new Thickness(0) };
            StackPanel sp = new() { Margin = new Thickness(5) };
            sp.Width = 60;
            ContextMenu mn = new();

            MenuItem miE = new() { Header = "Edit" };
            MenuItem miM = new() { Header = "Move.." };
            MenuItem miRemove = new() { Header = "Remove" };

            mn.Opened += (e, s) => {
                miE.IsEnabled = !isLocked;
                miM.IsEnabled = !isLocked;
                miRemove.IsEnabled = !isLocked;
            };

            miRemove.Click += (sender, e) =>
            {
                var items = geticons();
                if (isLocked || items == null) return;
                items.Remove(icon);
                wpcont.Children.Remove(btn);
                App.save();
            };

            mn.Items.Add(miE);
            mn.Items.Add(miM);
            mn.Items.Add(miRemove);
            btn.ContextMenu = mn;
            Image ico = new() { Width = 36, Height = 36, Margin = new Thickness(9) };
            
            sp.Children.Add(ico);
            TextBlock lbl = new() { TextWrapping = TextWrapping.Wrap, TextTrimming = TextTrimming.CharacterEllipsis, HorizontalAlignment = HorizontalAlignment.Center, Foreground = Brushes.White };
            lbl.MaxHeight = (lbl.FontSize * 1.5) + (lbl.Margin.Top * 2);
            void updateNameAndIcon()
            {
                if (icon.DisplayName == "{AUTONAME}")
                {
                    lbl.Text = Path.GetFileNameWithoutExtension(icon.Filename);
                }
                else
                {
                    lbl.Text = icon.DisplayName;
                }
                try
                {
                    if (icon.DisplayIcon == "{AUTOICON}")
                    {
                        if (Directory.Exists(icon.Filename)) //is it a folder?
                        {
                            ico.Source = new BitmapImage(new Uri("pack://application:,,,/folder-White.png"));
                        }
                        else
                        {
                            var extractedIcon = System.Drawing.Icon.ExtractAssociatedIcon(icon.Filename);
                            if (extractedIcon != null)
                            {
                                ico.Source = extractedIcon.ToImageSource();
                            }
                        }
                    }
                    else
                    {
                        ico.Source = new BitmapImage(new Uri(icon.DisplayIcon, UriKind.Relative));
                    }
                }
                catch
                { }
            }
            updateNameAndIcon();
            sp.Children.Add(lbl);
            miM.Click += (sender, e) => {
                var items = geticons();
                if (isLocked || items == null) return;
                StackPanel cnt = new();
                Window wwin = new() { Title = "Move " + icon.DisplayName, Content = cnt, Width = 300, Height = 100, WindowStartupLocation = WindowStartupLocation.CenterScreen, ResizeMode = ResizeMode.NoResize };
                ComboBox lv = new();
                foreach (FenceItem icn in items)
                {
                    //StackPanel cc = new() { Orientation = Orientation.Horizontal};
                    //cc.Children.Add(new Image() { Source = ico.Source });
                    //cc.Children.Add(new Label() { Content = lbl.Text });
                    lv.Items.Add(icn.Filename);
                }
                cnt.Children.Add(lv);
                Button mbtn = new() { Content = "Move" };
                cnt.Children.Add(mbtn);
                mbtn.Click += (sender, e) => {
                    if (lv.SelectedIndex == -1)
                    {
                        return;
                    }
                    int id = wpcont.Children.IndexOf(btn);
                    FenceItem olddata = items[lv.SelectedIndex];
                    items[lv.SelectedIndex] = items[id];
                    items[id] = olddata;
                    App.save();
                    initcontent();
                    wwin.Close();
                };
                wwin.ShowDialog();
            };
            miE.Click += (sender, e) => {
                var items = geticons();
                if (isLocked || items == null)
                {
                    return;
                }


                StackPanel cnt = new();
                Window wwin = new() { Title = "Edit " + icon.Filename, Content = cnt, Width = 450, Height = 200, WindowStartupLocation = WindowStartupLocation.CenterScreen, ResizeMode = ResizeMode.NoResize };
                TextBox createsec(string name, string defaulval)
                {
                    DockPanel dpp = new();
                    Label lbl = new() { Content = name };
                    dpp.Children.Add(lbl);
                    TextBox tbb = new() { Text = defaulval };
                    dpp.Children.Add(tbb);
                    cnt.Children.Add(dpp);
                    return tbb;
                }
                ;
                int id = wpcont.Children.IndexOf(btn);
                TextBox tbDN = createsec("Display Name", items[id].DisplayName);
                TextBox tbDI = createsec("Display Icon", items[id].DisplayIcon);
                Button abtn = new() { Content = "Apply" };
                abtn.Click += (sender, e) => {
                    items[id].DisplayIcon = tbDI.Text;
                    items[id].DisplayName = tbDN.Text;
                    updateNameAndIcon();
                    App.save();
                    wwin.Close();
                };
                cnt.Children.Add(abtn);
                wwin.ShowDialog();
            };
            var p = new Process();
            p.StartInfo = new ProcessStartInfo(icon.Filename)
            {
                UseShellExecute = true
            };
            btn.Click += (sender, e) => {
                try { p.Start(); } catch (Exception ex) { MessageBox.Show(ex.ToString(), "Error while launching", MessageBoxButton.OK, MessageBoxImage.Error); }
            };
            btn.Content = sp;
            wpcont.Children.Add(btn);
        }

        public void applyFenceSettings()
        {
            win.ResizeMode = isLocked ? ResizeMode.NoResize : ResizeMode.CanResize;
            win.Title = Title;
            titlelabel.Content = Title;
            cborder.Background = new SolidColorBrush(Color.FromArgb(fencecolor[0], fencecolor[1], fencecolor[2], fencecolor[3]));
        }

        public List<FenceItem>? geticons()
        {
            if (!(Items is List<FenceItem>) && ItemsType != "Data")
            {
                return null;
            }
            return Items as List<FenceItem>;
        }
    }

    public class FenceItem
    {
        public string Filename = "";
        public string DisplayName = "{AUTONAME}";
        public string DisplayIcon = "{AUTOICON}";
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
                throw new Win32Exception("Couldn't delete the bitmap object!");
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
                    SetCurrentProcessExplicitAppUserModelID("Birdy.Fences");

                    // Set TOOLWINDOW style to hide from Alt+Tab
                    var exStyle = (int)GetWindowLongPtr(hWnd, GWL_EXSTYLE);
                    exStyle |= WS_EX_TOOLWINDOW;
                    SetWindowLongPtr(hWnd, GWL_EXSTYLE, (IntPtr)exStyle);
                };
            }
        }
    }
}
