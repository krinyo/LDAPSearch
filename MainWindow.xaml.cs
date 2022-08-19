using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LDAPSearch
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region ResizeWindows
        bool ResizeInProcess = false;
        private void Resize_Init(object sender, MouseButtonEventArgs e)
        {
            Rectangle senderRect = sender as Rectangle;
            if (senderRect != null)
            {
                ResizeInProcess = true;
                senderRect.CaptureMouse();
            }
        }

        private void Resize_End(object sender, MouseButtonEventArgs e)
        {
            Rectangle senderRect = sender as Rectangle;
            if (senderRect != null)
            {
                ResizeInProcess = false; ;
                senderRect.ReleaseMouseCapture();
            }
        }

        private void Resizeing_Form(object sender, MouseEventArgs e)
        {
            if (ResizeInProcess)
            {
                Rectangle senderRect = sender as Rectangle;
                Window mainWindow = senderRect.Tag as Window;
                if (senderRect != null)
                {
                    double width = e.GetPosition(mainWindow).X;
                    double height = e.GetPosition(mainWindow).Y;
                    senderRect.CaptureMouse();
                    if (senderRect.Name.ToLower().Contains("right"))
                    {
                        width += 5;
                        if (width > 0)
                            mainWindow.Width = width;
                    }
                    if (senderRect.Name.ToLower().Contains("left"))
                    {
                        width -= 5;
                        mainWindow.Left += width;
                        width = mainWindow.Width - width;
                        if (width > 0)
                        {
                            mainWindow.Width = width;
                        }
                    }
                    if (senderRect.Name.ToLower().Contains("bottom"))
                    {
                        height += 5;
                        if (height > 0)
                            mainWindow.Height = height;
                    }
                    if (senderRect.Name.ToLower().Contains("top"))
                    {
                        height -= 5;
                        mainWindow.Top += height;
                        height = mainWindow.Height - height;
                        if (height > 0)
                        {
                            mainWindow.Height = height;
                        }
                    }
                }
            }
        }
        #endregion
        #region TitleButtons
        private void MinimizeWindow(object sender, RoutedEventArgs e)
        {
            App.Current.MainWindow.WindowState = WindowState.Minimized;
        }

        private void MaximizeClick(object sender, RoutedEventArgs e)
        {
            AdjustWindowSize();
        }

        private void StackPanel_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                if (e.ClickCount == 2)
                {
                    AdjustWindowSize();
                }
                else
                {
                    App.Current.MainWindow.DragMove();
                }
            }
        }

        private void AdjustWindowSize()
        {
            if (App.Current.MainWindow.WindowState == WindowState.Maximized)
            {
                App.Current.MainWindow.WindowState = WindowState.Normal;
                MaximizeButton.Content = "";
            }
            else if (App.Current.MainWindow.WindowState == WindowState.Normal)
            {
                App.Current.MainWindow.WindowState = WindowState.Maximized;
                MaximizeButton.Content = "";
            }
        }


        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            App.Current.MainWindow.Close();
        }
        #endregion
        /*LDAP functions*/
        private string GetCurrentDomainPath()
        {
            DirectoryEntry de = new DirectoryEntry("LDAP://RootDSE");

            return "LDAP://" + Prefix + de.Properties["defaultNamingContext"][0].ToString();
        }
        private string FormSearchFilter(string crit) 
        {
            string result = "(| ";
            foreach (var prop in UserProps) 
            {
                if(prop.Value != "distinguishedName")
                    result += "("+prop.Value+"="+crit+"*)";
            }
            return result += ")";
        }
        private void GetAdditionalUserInfo(List<Dictionary<string, string>> userRecords)
        {
            SearchResultCollection results;
            DirectorySearcher ds = null;
            DirectoryEntry de = new DirectoryEntry(GetCurrentDomainPath());

            ds = new DirectorySearcher(de);

            //Loading properties
            foreach (var el in UserProps) 
            {
                ds.PropertiesToLoad.Add(el.Value);
            }
            ds.Filter = "(&(objectCategory=User)(objectClass=person))";

            results = ds.FindAll();

            foreach (SearchResult sr in results)
            {
                var bufListItem = new Dictionary<string, string>();
                userRecords.Add(bufListItem);
                foreach (var prop in UserProps) {
                    if (sr.Properties[prop.Value].Count > 0) {
                        if (prop.Value == "distinguishedName")
                        {
                            if (sr.Properties[prop.Value][0].ToString().Split(",")[1] == "OU=Disabled Accounts")
                                bufListItem.Add(prop.Key, "Disabled");
                            else
                                bufListItem.Add(prop.Key, "Enabled");
                        }
                        else
                            bufListItem.Add(prop.Key, sr.Properties[prop.Value][0].ToString());
                    }
                }

            }

        }

        private void ShowAllUsersInfo() 
        {
            ClearButtons();
            RecPanel.Children.Clear();
            foreach (var rec in UserRecords) 
            {
                TextBlock strText;
                strText = new TextBlock();
                strText.FontSize = 14;
                strText.Foreground = Brushes.Bisque;
                bool hasLogin = false;
                string pcName = "";
                foreach (var key in rec.Keys) 
                {
                    strText.Text += key + " " + rec[key] + "\n";
                    if (key == "Login info") 
                    {
                        hasLogin = true;
                        pcName = rec[key].Split(" ")[4];
                    }
                }
                strText.Text += "";
                RecPanel.Children.Add(strText);

                if (hasLogin)
                {
                    RecPanel.Children.Add(new TextBlock() { Text = "Radmin control", Foreground = Brushes.Bisque});
                    Button btn = new Button();
                    //ButtonNames.Add(pcName, btn);
                    btn.Width = 150;
                    btn.Height = 20;
                    btn.Content = pcName;
                    btn.HorizontalAlignment = HorizontalAlignment.Left;
                    btn.VerticalAlignment = VerticalAlignment.Top;
                    btn.Click += new RoutedEventHandler(Radmin_OnClick);
                    
                    RecPanel.Children.Add(btn);

                    RecPanel.Children.Add(new TextBlock() { Text = "Radmin No Control", Foreground = Brushes.Bisque });
                    Button nctrlbtn = new Button();
                    //ButtonNames.Add(pcName, nctrlbtn);
                    nctrlbtn.Width = 150;
                    nctrlbtn.Height = 20;
                    nctrlbtn.Content = pcName;
                    nctrlbtn.HorizontalAlignment = HorizontalAlignment.Left;
                    nctrlbtn.VerticalAlignment = VerticalAlignment.Top;
                    nctrlbtn.Click += new RoutedEventHandler(Radmin_No_Control_OnClick);

                    RecPanel.Children.Add(nctrlbtn);
                }
                RecPanel.Children.Add(new TextBlock() { Text = "_______________", Foreground = Brushes.Bisque});
            }

            
        }
        private DirectorySearcher BuildUserSearcher(DirectoryEntry de)
        {
            DirectorySearcher ds = null;

            ds = new DirectorySearcher(de);
            foreach (var prop in UserProps)
            {
                ds.PropertiesToLoad.Add(prop.Value);
            }
            return ds;
        }
        private void ClearUsers() 
        {
            UserRecords.Clear();
        }
        private void ClearButtons() 
        {
            ButtonNames.Clear();
        }
        private void SearchForUsers(string criterion, List<Dictionary<string, string>> userRecords)
        {
            SearchResultCollection results;
            DirectorySearcher ds = null;
            DirectoryEntry de = new DirectoryEntry(GetCurrentDomainPath());

            // Build User Searcher
            ds = BuildUserSearcher(de);

            //ds.Filter = "(&(objectCategory=User)(objectClass=person)(name=" + criterion + "*))";
            ds.Filter = FormSearchFilter(criterion);
            results = ds.FindAll();
            ClearUsers();

            foreach (SearchResult sr in results)
            {
                //nm.Text += (sr.GetPropertyValue("cn") + " \n");

                var bufListItem = new Dictionary<string, string>();
                userRecords.Add(bufListItem);
                foreach (var prop in UserProps)
                {
                    if (sr.Properties[prop.Value].Count > 0)
                    {
                        if (prop.Value == "distinguishedName")
                        {
                            if (sr.Properties[prop.Value][0].ToString().Split(",")[1] == "OU=Disabled Accounts")
                                bufListItem.Add(prop.Key, "Disabled");
                            else
                                bufListItem.Add(prop.Key, "Enabled");
                        }
                        else
                            bufListItem.Add(prop.Key, sr.Properties[prop.Value][0].ToString());
                    }
                }
            }
        }
        /*              */
        private string Title = "LDAPSearch";
        private string Prefix = "OU=SibirCentrUsers,";
        private List<Dictionary<string, string>> UserRecords = new List<Dictionary<string, string>>();
        private Dictionary<string, Button> ButtonNames = new Dictionary<string, Button>();
        private string RadminStartCommand = @"C:\Program Files (x86)\Radmin Viewer 3\Radmin.exe";
        private string Cmd = "cmd.exe";
        private Dictionary<string, string> UserProps = new Dictionary<string, string>()
        {
            { "Name", "cn" },
            { "Login", "sAMAccountName" },
            
            { "Account availibility", "distinguishedName"},

            { "Tel", "telephoneNumber" },
            { "Mob", "mobile" },
            { "Position", "title" },
            { "Dep", "departament" },
            { "Login info", "info" }
        };

        public MainWindow()
        {
            InitializeComponent();
            HeaderTitle.Text = Title;
            //GetAdditionalUserInfo(UserRecords);
            //ShowAllUsersInfo();

        }

        private void Search_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(Search.Text.Length > 2)
                SearchForUsers(Search.Text, UserRecords);
                ShowAllUsersInfo();
        }

        private void Radmin_OnClick(object sender, RoutedEventArgs e)
        {
            string command = RadminStartCommand;
            var proc = new ProcessStartInfo( RadminStartCommand,  "/connect:" + ((Button)sender).Content);


            Process.Start(proc);
        }
        private void Radmin_No_Control_OnClick(object sender, RoutedEventArgs e)
        {
            string command = RadminStartCommand;
            var proc = new ProcessStartInfo(RadminStartCommand, "/connect:" + ((Button)sender).Content + " /noinput");


            Process.Start(proc);
        }
    }
}
