using Microsoft.Win32;
using Newtonsoft.Json;
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
        public static string userdir = "";
        public static List<Fence> fencedata = new();

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, out SHFileInfo psfi, uint cbFileInfo, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DestroyIcon(IntPtr hIcon);

        public const uint SHGFI_ICON = 0x000000100;
        public const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
        public const uint SHGFI_OPENICON = 0x000000002;
        public const uint SHGFI_SMALLICON = 0x000000001;
        public const uint SHGFI_LARGEICON = 0x000000000;
        public const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;
        public const uint FILE_ATTRIBUTE_FILE = 0x00000100;
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHFileInfo
        {
            public IntPtr hIcon;

            public int iIcon;

            public uint dwAttributes;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        };

        public enum IconSize : short
        {
            Small,
            Large
        }

        public enum ItemState : short
        {
            Undefined,
            Open,
            Close
        }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            userdir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            Directory.CreateDirectory(userdir + "\\Birdy Fences");
            Directory.CreateDirectory(userdir + "\\Birdy Fences\\Shortcuts");
            // fences.json exist but contains only []
            if (!File.Exists(userdir + "\\Birdy Fences\\fences.json") || File.ReadAllText(userdir + "\\Birdy Fences\\fences.json").Trim() == "[]")
            {
                List<FenceItem> defconfItems = [
                    new FenceItem() {
                        Filename = userdir + "\\Birdy Fences\\WELCOME.txt",
                        DisplayName = "Welcome!"
                    }
                ];
                List<Fence> defconf =
                [
                    new Fence()
                    {
                        Title = "Welcome to BirdyFences!",
                        Items = defconfItems
                    },
                ];
                File.WriteAllText(userdir + "\\Birdy Fences\\fences.json", JsonConvert.SerializeObject(defconf));
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
            fencedata = JsonConvert.DeserializeObject<List<Fence>>(File.ReadAllText(userdir + "\\Birdy Fences\\fences.json")) ?? new();
            if (fencedata == null) fencedata = new();
            
            if (fencedata != null)
            {
                foreach (Fence fence in fencedata)
                {
                    fence.init();
                }
            }
        }

        public static async void save()
        {
            await File.WriteAllTextAsync(userdir + "\\Birdy Fences\\fences.json", JsonConvert.SerializeObject(fencedata));
        }

        public static BitmapSource GetIcon(string path, IconSize size, ItemState state)
        {
            var flags = (uint)(SHGFI_ICON | SHGFI_USEFILEATTRIBUTES);
            var attribute = (uint)(FILE_ATTRIBUTE_FILE);

            if (object.Equals(size, IconSize.Small))
            {
                flags += SHGFI_SMALLICON;
            }
            else
            {
                flags += SHGFI_LARGEICON;
            }
            var shfi = new SHFileInfo();
            var res = SHGetFileInfo(path, attribute, out shfi, (uint)Marshal.SizeOf(shfi), flags);
            if (object.Equals(res, IntPtr.Zero)) throw Marshal.GetExceptionForHR(Marshal.GetHRForLastWin32Error()) ?? new Exception("Unknown exception occurred!");
            try
            {
                var i = Imaging.CreateBitmapSourceFromHIcon(
                            shfi.hIcon,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                return i;
            }
            catch
            {
                throw;
            }
            finally
            {
                DestroyIcon(shfi.hIcon);
            }
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
        MenuItem miNewFence = new() { Header = "New Fence" };
        MenuItem miNewPortal = new() { Header = "New Portal Fence" };
        MenuItem miRemove = new() { Header = "Remove Fence" };
        MenuItem miColor = new() { Header = "Color" };
        MenuItem miLock = new() { Header = "Lock Fence", IsCheckable = true };
        Window win = new() { AllowDrop = true, AllowsTransparency = true, Background = Brushes.Transparent, ShowInTaskbar = false, WindowStyle = WindowStyle.None };
        Border titleborder = new() { Background = new SolidColorBrush(Color.FromArgb(20, 0, 0, 0)), CornerRadius = new CornerRadius(6) };
        DockPanel titlecont = new();
        Label titlelabel = new() { Foreground = Brushes.White, HorizontalContentAlignment = HorizontalAlignment.Center };
        TextBox titletb = new() { HorizontalContentAlignment = HorizontalAlignment.Center, Visibility = Visibility.Collapsed, Background = Brushes.Transparent, Foreground = Brushes.White, Padding = new Thickness(4) };
        WrapPanel wpcont = new() { AllowDrop = true };
        ScrollViewer wpcontscr = new() { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };

        public void init()
        {
            bool dragging = false;
            Point? lastmousepos = null;

            //Menu
            cm.Items.Add(miNewFence);
            cm.Items.Add(miNewPortal);
            cm.Items.Add(miRemove);
            cm.Items.Add(new Separator());
            cm.Items.Add(miColor);
            cm.Items.Add(miLock);

            // Title stuff
            Grid namecont = new();
            namecont.Children.Add(titlelabel);
            namecont.Children.Add(titletb);
            titlecont.Children.Add(namecont);
            dp.Children.Add(titleborder);
            dp.Children.Add(wpcontscr);

            // Window stuff
            HideAltTab.VirtualDesktopHelper.MakeWindowPersistent(win);
            WindowChrome.SetWindowChrome(win, new WindowChrome() { CaptionHeight = 0, ResizeBorderThickness = new Thickness(5) });
            DockPanel.SetDock(titleborder, Dock.Top);

            // Other argument stuff
            win.Width = Width;
            win.Height = Height;
            win.Top = Y;
            win.Left = X;

            win.Content = cborder;
            win.ContextMenu = cm;
            cborder.Child = dp;
            titleborder.Child = titlecont;
            wpcontscr.Content = wpcont;


            win.KeyDown += (s, e) => { 
                if (e.Key == Key.System && e.SystemKey == Key.F4) // Use this to dedect alt+f4 because Closing would fire too on fence deletion.
                {
                    App.save();
                    Application.Current.Shutdown();
                }
            };

            bool makefront = false;
            win.Activated += (s, e) => {
                if (makefront) //Ignore first one because it happens during init in a foreach.
                {
                    //Move to last so it creates last which makes it at front
                    App.fencedata.Remove(this);
                    App.fencedata.Add(this);
                    App.save();
                }
                makefront = true;
            };

            cm.Opened += (e,s) => {
                miLock.IsChecked = isLocked;
                miRemove.IsEnabled = !isLocked;
            };

            miLock.Click += (sender, e) =>
            {
                // Toggle fence lock: disables/enables resizing the fence
                isLocked = !isLocked;
                applyFenceSettings();
                App.save();
            };
            miColor.Click += (s, e) => {
                // Color picker
                StackPanel mcont = new();
                Window win = new() { Title = "Choose color", ResizeMode = ResizeMode.NoResize, Width = 300, Height = 150, WindowStartupLocation = WindowStartupLocation.CenterScreen, Content = mcont };
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

            miRemove.Click += (sender, e) => {
                // Remove fence
                if (!isLocked)
                {
                    App.fencedata.Remove(this);
                    win.Close();
                    App.save();
                    cm.Items.Refresh();
                }
            };

            miNewFence.Click += (sender, e) => {
                // New fence
                var fnc = new Fence();
                App.fencedata.Add(fnc);
                fnc.init();
                App.save();
            };

            miNewPortal.Click += (sender, e) => {
                // New portal fence
                var dialog = new OpenFolderDialog
                {
                    Title = "Select folder for portal",
                    ValidateNames = true,
                };
                
                if (dialog.ShowDialog() == true)
                {
                    var fnc = new Fence()
                    {
                        ItemsType = "Portal",
                        Items = dialog.FolderName,
                        Title = Path.GetFileNameWithoutExtension(dialog.FolderName)
                    };
                    App.fencedata.Add(fnc);
                    fnc.init();
                    App.save();
                }
            };

            titlelabel.MouseDown += (object sender, MouseButtonEventArgs e) => {
                if (e.ClickCount == 1)
                {
                    if (e.LeftButton == MouseButtonState.Pressed)
                    {
                        //win.DragMove();
                        dragging = true;
                        win.CaptureMouse();
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
            win.MouseUp += (object sender, MouseButtonEventArgs e) => {
                if (dragging)
                {
                    dragging = false;
                    lastmousepos = null;
                    win.ReleaseMouseCapture();
                    Y = win.Top;
                    X = win.Left;
                    App.save();
                }
            };


            
            win.MouseMove += (object sender, MouseEventArgs e) => {
                if (dragging && !isLocked)
                {
                    Point pos = win.PointToScreen(e.GetPosition(win));
                    if (lastmousepos != null)
                    {
                        var lastpos = (Point)lastmousepos;
                        double diffx = pos.X - lastpos.X;
                        double diffy = pos.Y - lastpos.Y;
                        double newx = win.Left + diffx;
                        double newy = win.Top + diffy;
                        bool snappedX = false;
                        bool snappedY = false;
                        foreach (Fence f in App.fencedata)
                        {
                            if (f == this)
                            {
                                continue;
                            }

                            bool snapx = Math.Floor(f.X / 10) == Math.Floor(newx / 10);
                            bool snapy = Math.Floor(f.Y / 10) == Math.Floor(newy / 10);
                            if (snapx)
                            {
                                newx = f.X;
                                snappedX = true;
                            }
                            if (snapy)
                            {
                                newy = f.Y;
                                snappedY = true;
                            }
                        }
                        if (!snappedX && !snappedY)
                        {
                            lastmousepos = pos;
                        }else if (snappedX)
                        {
                            lastmousepos = new Point(lastpos.X, pos.Y);
                        }
                        else if (snappedY)
                        {
                            lastmousepos = new Point(pos.X, lastpos.Y);
                        }
                        win.Top = newy;
                        win.Left = newx;
                    }else
                        lastmousepos = pos;
                }
                
            };

            win.SizeChanged += (sender, e) => {
                Width = win.ActualWidth;
                Height = win.ActualHeight;
                Y = win.Top;
                X = win.Left;
                App.save();
            };
            
            // Dropping (To add icons)
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
                try
                {
                    string[] dirs = Directory.GetDirectories(dpath);
                    foreach (string dir in dirs)
                    {
                        if (Path.GetFileName(dir).StartsWith(".")) continue;
                        FenceItem icon = new FenceItem() { Filename = dir };
                        addicon(icon);
                    }
                    string[] files = Directory.GetFiles(dpath);
                    foreach (string file in files)
                    {
                        if (Path.GetFileName(file).ToLower() == "desktop.ini") continue;
                        if (Path.GetFileName(file).StartsWith(".")) continue;
                        FenceItem icon = new FenceItem() { Filename = file };
                        addicon(icon);
                    }
                }catch //(Exception e) TODO: maybe add a indicator, messagebox will crash it rn
                {
                    //MessageBox.Show(e.ToString(), "Couldn't load fence", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        void addicon(FenceItem icon)
        {
            Button btn = new() { Margin = new Thickness(5), Style = (Style)Application.Current.Resources["asbsLight"], Background = Brushes.Transparent, BorderThickness = new Thickness(0) };
            StackPanel sp = new() { Margin = new Thickness(5) };
            sp.Width = 60;
            ContextMenu mn = new();

            MenuItem miEdit = new() { Header = "Edit" };
            MenuItem miMove = new() { Header = "Swap with..." };
            MenuItem miRemove = new() { Header = "Remove" };

            mn.Opened += (e, s) => {
                var geticonsstatus = geticons() != null;
                miEdit.IsEnabled = !isLocked && geticonsstatus;
                miMove.IsEnabled = !isLocked && geticonsstatus;
                miRemove.IsEnabled = !isLocked && geticonsstatus;
            };

            miRemove.Click += (sender, e) =>
            {
                var items = geticons();
                if (isLocked || items == null) return;
                items.Remove(icon);
                wpcont.Children.Remove(btn);
                App.save();
            };

            mn.Items.Add(miEdit);
            mn.Items.Add(miMove);
            mn.Items.Add(miRemove);
            btn.ContextMenu = mn;
            Image ico = new() { Width = 36, Height = 36, Margin = new Thickness(9) };
            
            sp.Children.Add(ico);
            TextBlock lbl = new() { TextWrapping = TextWrapping.Wrap, TextTrimming = TextTrimming.CharacterEllipsis, HorizontalAlignment = HorizontalAlignment.Center, Foreground = Brushes.White };
            lbl.MaxHeight = (lbl.FontSize * 1.5) + (lbl.Margin.Top * 2);
            void updateNameAndIcon()
            {
                lbl.Text = icon.getDisplayName();
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
                            var extractedIcon = App.GetIcon(icon.Filename, App.IconSize.Large, App.ItemState.Undefined);
                            ico.Source = extractedIcon;
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
            miMove.Click += (sender, e) => {
                var items = geticons();
                if (isLocked || items == null) return;
                StackPanel cnt = new();
                Window wwin = new() { Title = "Move " + icon.getDisplayName(), Content = cnt, Width = 300, Height = 80, WindowStartupLocation = WindowStartupLocation.CenterScreen, ResizeMode = ResizeMode.NoResize };
                ComboBox lv = new();
                foreach (FenceItem icn in items)
                {
                    //StackPanel cc = new() { Orientation = Orientation.Horizontal};
                    //cc.Children.Add(new Image() { Source = ico.Source });
                    //cc.Children.Add(new Label() { Content = lbl.Text });
                    lv.Items.Add(icn.getDisplayName());
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
            miEdit.Click += (sender, e) => {
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

        public string getDisplayName()
        {
            if (DisplayName == "{AUTONAME}")
            {
                return Path.GetFileNameWithoutExtension(Filename);
            }
            else
            {
                return DisplayName;
            }
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
