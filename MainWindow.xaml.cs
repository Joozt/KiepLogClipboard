using System.Windows;
using System;
using WindowsInput;
using System.Timers;
using System.Windows.Threading;
using System.Globalization;
using System.Threading;

namespace KiepLogClipboard
{
    public partial class MainWindow : Window
    {
        private string clipboardText;

        public MainWindow()
        {
            InitializeComponent();
            progressBar.Value = 0;

           

            // Start progress bar
            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(100);
            timer.Tick += new EventHandler(timer_Tick);
            timer.Start();


            if (!Clipboard.ContainsText())
            {
                Console.WriteLine("clipboard is empty");
                Application.Current.Shutdown();
            }

            clipboardText = Clipboard.GetText();
            if (clipboardText == null || clipboardText.Length == 0)
            {
                Console.WriteLine("clipboard is empty");
                Application.Current.Shutdown();
            }


            Thread myThread = new Thread(processDatesAndGetMails); // <-- remember, don't put the () after the method you want running on that thread!
            myThread.Start();



            //Juni 2012, 4 - juni 2012, 8

        }

        void timer_Tick(object sender, EventArgs e)
        {
            progressBar.Value += 4;
            if (progressBar.Value >= 100)
            {
                progressBar.Value = 0;
            }
        }

        private void processDatesAndGetMails()
        {


            string[] dates = clipboardText.Split('-');
            if (dates.Length != 2)
            {
                Console.WriteLine("dates does not contain 2 items");
                return;
            }

            string fromDateString = dates[0].Trim();
            if (fromDateString == null || fromDateString.Length == 0)
            {
                Console.WriteLine("fromDateString is empty");
                return;
            }

            string toDateString = dates[1].Trim();
            if (toDateString == null || toDateString.Length == 0)
            {
                Console.WriteLine("toDateString is empty");
                return;
            }

            DateTime fromDate = parseDate(fromDateString);
            if (fromDate.Equals(DateTime.MinValue))
            {
                Console.WriteLine("fromDate is empty");
                return;
            }

            DateTime toDate = parseDate(toDateString);
            if (toDate.Equals(DateTime.MinValue))
            {
                Console.WriteLine("toDate is empty");
                return;
            }

            KiepMail kiepMail = new KiepMail();
            string output = kiepMail.Download(fromDate, toDate);
            sendToTobii(output);

            Application.Current.Shutdown();
        }

        private void sendToTobii(string output)
        {
            Clipboard.SetText(output);
            InputSimulator inputSimulator = new InputSimulator();
            inputSimulator.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.NUMPAD1);
        }

        private DateTime parseDate(string dateString)
        {
            CultureInfo MyCultureInfo = new CultureInfo("nl-NL");
            try
            {
                DateTime date = DateTime.ParseExact(dateString, "MMMM yyyy, d", MyCultureInfo);
                Console.WriteLine(date);
                return date;
            }
            catch (FormatException)
            {
                Console.WriteLine("Unable to parse '{0}'", dateString);
                return DateTime.MinValue;
            }
        }
    }
}
