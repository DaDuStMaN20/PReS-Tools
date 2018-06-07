using System;
using System.Collections.Generic;
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
using NHotkey.Wpf;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using WindowsInput;
using System.Windows.Interop;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace PReS_Programmer_Utility_Panel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        System.Collections.Generic.List<String> actualClipboard;
        bool justWroteToClipboard = false;
        String currentItem;
        
        public MainWindow()
        {
            InitializeComponent();
            actualClipboard = new List<String>();
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            // Initialize the clipboard now that we have a window soruce to use
            var windowClipboardManager = new ClipboardManager(this);
            windowClipboardManager.ClipboardChanged += ClipboardChanged;
        }
        //Clipboard
        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            //get the index of the selected item
            int selectedIndex = ClipboardListBox.SelectedIndex;

            if (selectedIndex != -1)
            {
               
                //copy to clipboard
                
                System.Windows.Clipboard.SetText(actualClipboard[selectedIndex].ToString());
                //move the item up to the top of the list using the method
                MoveToTop(selectedIndex);
                //rebuild the list
                Repopulate();
                currentItem = actualClipboard[selectedIndex].ToString();
            }

           
            justWroteToClipboard = true;

        }

        private void ClipboardChanged(object sender, EventArgs e)
        {
            String clipboardText;
            try
            {

                // Handle your clipboard update here, debug logging example:
                if (Clipboard.ContainsText())
                {
                
                    clipboardText = Clipboard.GetText(TextDataFormat.UnicodeText);
                    if(!clipboardText.Equals(currentItem)){ //make sure that if you click an item, it can differentiate from the item that was copied
                                                            // it goes into this multiple times after writing to clipboard, so that is how I stopped it
                        if (!actualClipboard.Contains(clipboardText))
                        { //if it has the thing already
                        
                            actualClipboard.Insert(0, clipboardText);

                            if(actualClipboard.Count > 10)
                            {
                                actualClipboard.RemoveAt(10);
                            }

                        }
                        else
                        {
                            MoveToTop(actualClipboard.IndexOf(clipboardText));
                        }

                        Repopulate();
                    }
                
                } else
                {
                    justWroteToClipboard = false;
                }

            }
            catch (Exception err)
            {
                //do nothing
            }
        }

        public void Repopulate()
        {
            String listText;

            //this!
            ClipboardListBox.Items.Clear();

            foreach (String s1 in actualClipboard)
            {
                listText = s1;

                if (s1.Length > 100)
                {
                    listText = s1.Substring(0, 100);

                    listText += " ...";
                }
                else
                {

                    listText = s1;
                }

                listText = listText.Replace("\r\n", " ");

                ClipboardListBox.Items.Add(listText);
            }
            
        }

        public void MoveToTop(int index)
        {
            var item = actualClipboard[index];
            actualClipboard.RemoveAt(index);
            actualClipboard.Insert(0, item);
            
        }

        public void MoveToBottom(ListBox lb, int index)
        {
            var item = lb.Items[index];
            lb.Items.RemoveAt(index);
            lb.Items.Add(item);
            //lb.Refresh();
        }

        //Shortcuts
        public void NewPracticeHandler(object sender, NHotkey.HotkeyEventArgs e)
        {
            SendTextOut(";ITR DW Added New Practice DATE ");
        }

        public void ProgHandler(object sender, NHotkey.HotkeyEventArgs e)
        {
            SendTextOut("_prog(ITR)_");
        }

        public void NewLogoHandler(object sender, NHotkey.HotkeyEventArgs e)
        {
            SendTextOut(";ITR DW Added Logo DATE ");
        }

        public void ITRHandler(object sender, NHotkey.HotkeyEventArgs e)
        {
            SendTextOut("ITR");
        }

        public void ResetShortcut(String shortcutName)
        {
            HotkeyManager.Current.AddOrReplace(shortcutName, Key.NoName, ModifierKeys.Control, ResetHandler);
        }

        public void JustinCommentHandler(object sender, NHotkey.HotkeyEventArgs e)
        {
            SendTextOut(";ITR jmccabe DATE ");
        }

        public void ResetHandler(object sender, NHotkey.HotkeyEventArgs e)
        {
            //do nothing
        }

        //Setting a shortcut (shortcut name, ctrl?, numpad key?, text to go into the comment, include itr?, include date?)
        public void SendTextOut(String text)
        {
            String sysDate = "";               //System Date
            DateTime newDate = DateTime.Now;   //date time to be converted to string

            String ITRnum = "";                //ITR number
            String windowTitle = "";           //Title of the PReS IDE window
            int itrStartPos;                   //Start of ITRnum

            String shortcutString = text;      //shortcut text
            var sim = new InputSimulator();    //Simulates the keyboard

            
            //if ITR is true, get ITR
            if (shortcutString.Contains("ITR"))
            {

                Process[] processlist = Process.GetProcesses();
               
                //get process names
                foreach (Process process in processlist)
                {
                    if (!String.IsNullOrEmpty(process.MainWindowTitle))
                    {
                        //search for pres ide
                        if (process.MainWindowTitle.Contains("PReS Integrated Development Environment"))
                        {
                            windowTitle = process.MainWindowTitle;
                            break; //We dont care about anything else after the pres window
                        }
                    }
                }

                //Parse the ITR out of the string
                itrStartPos = windowTitle.IndexOf("_prog(")+6;      //Find the start position of the ITR
                ITRnum = windowTitle.Substring(itrStartPos, 6);     //Create the substring by taking the start pos and the length of the ITR

                shortcutString = shortcutString.Replace("ITR", ITRnum);              //Replace any instance of ITR in the shortcut text with the actual ITR
                
            }

            //IF date is true, put get the system date
            if (shortcutString.Contains("DATE"))
            {
                sysDate = newDate.ToString("MM/dd/yy"); //convert the date to the correct format

                shortcutString = shortcutString.Replace("DATE", sysDate);
            }

            sim.Keyboard
                .TextEntry(shortcutString);           
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            HotkeyManager.Current.AddOrReplace("NewPractice", Key.NumPad1, ModifierKeys.Control, NewPracticeHandler);
        }

        private void CheckBox_Checked_1(object sender, RoutedEventArgs e)
        {
            HotkeyManager.Current.AddOrReplace("Prog", Key.NumPad0, ModifierKeys.Control, ProgHandler);
        }

        private void CheckBox_Checked_2(object sender, RoutedEventArgs e)
        {
            HotkeyManager.Current.AddOrReplace("NewLogo", Key.Add, ModifierKeys.Control, NewLogoHandler);
        }

        private void CheckBox_Checked_3(object sender, RoutedEventArgs e)
        {
            HotkeyManager.Current.AddOrReplace("ITR", Key.Decimal, ModifierKeys.Control, ITRHandler);
        }

        private void CheckBox_Checked_4(object sender, RoutedEventArgs e)
        {
            HotkeyManager.Current.AddOrReplace("JustinComment", Key.NumPad1, ModifierKeys.Control, JustinCommentHandler);
        }

        private void CheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            ResetShortcut("NewPractice");
        }

        private void CheckBox_UnChecked_1(object sender, RoutedEventArgs e)
        {
            ResetShortcut("Prog");
        }

        private void CheckBox_UnChecked_2(object sender, RoutedEventArgs e)
        {
            ResetShortcut("NewLogo");
        }

        private void CheckBox_UnChecked_3(object sender, RoutedEventArgs e)
        {
            ResetShortcut("ITR");
        }

        private void CheckBox_UnChecked_4(object sender, RoutedEventArgs e)
        {
            ResetShortcut("JustinComment");
        }

        private void List(object sender, MouseButtonEventArgs e)
        {

        }
    }

    internal static class NativeMethods
    {
        // See http://msdn.microsoft.com/en-us/library/ms649021%28v=vs.85%29.aspx
        public const int WM_CLIPBOARDUPDATE = 0x031D;
        public static IntPtr HWND_MESSAGE = new IntPtr(-3);

        // See http://msdn.microsoft.com/en-us/library/ms632599%28VS.85%29.aspx#message_only
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AddClipboardFormatListener(IntPtr hwnd);
    }

    public class ClipboardManager
    {
        public event EventHandler ClipboardChanged;

        public ClipboardManager(Window windowSource)
        {
            HwndSource source = PresentationSource.FromVisual(windowSource) as HwndSource;
            if (source == null)
            {
                throw new ArgumentException(
                    "Window source MUST be initialized first, such as in the Window's OnSourceInitialized handler."
                    , nameof(windowSource));
            }

            source.AddHook(WndProc);

            // get window handle for interop
            IntPtr windowHandle = new WindowInteropHelper(windowSource).Handle;

            // register for clipboard events
            NativeMethods.AddClipboardFormatListener(windowHandle);
        }

        private void OnClipboardChanged()
        {
            ClipboardChanged?.Invoke(this, EventArgs.Empty);
        }

        private static readonly IntPtr WndProcSuccess = IntPtr.Zero;

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == NativeMethods.WM_CLIPBOARDUPDATE)
            {
                OnClipboardChanged();
                handled = true;
            }

            return WndProcSuccess;
        }
    }
}


