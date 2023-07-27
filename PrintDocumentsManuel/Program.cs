using System.Diagnostics;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Printing;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Drawing;

namespace PrintDocumentsManuel
{
    internal class Program
    {
        private static CancellationTokenSource cancellationTokenSource = null;
        private static System.Drawing.Font printFont;
        private static StreamReader streamToPrint;

        static void Main(string[] args)
        {
            // Print PDF
            //cancellationTokenSource = new CancellationTokenSource();
            //cancellationTokenSource.CancelAfter(30000);
            //Console.WriteLine(Print_PDF(@"W:\Artikeldokumente\Vertrieb\Bedienungsanleitungen\02416.00\0241600D.PDF",
            //                            "IT Drucker", cancellationTokenSource.Token));

            // Print DOCX
            //cancellationTokenSource = new CancellationTokenSource();
            //cancellationTokenSource.CancelAfter(30000);
            //Console.WriteLine(Print_DOCX(@"C:\Users\homsi\Desktop\Files\test.docx"
            //            , "IT Drucker", cancellationTokenSource.Token));

            // Print Txt
            //cancellationTokenSource = new CancellationTokenSource();
            //cancellationTokenSource.CancelAfter(30000);
            //Print_TXT(@"C:\Users\MahmoudRahf\OneDrive\Desktop\Files\test.txt", "HP DeskJet 3630 series");
        }

        private static bool Print_PDF(string pPfad, string pDruckername, CancellationToken pCancellationToken)
        {
            // Ermitteln zur Laufzeit die Eigenschaften und die Typen dieser Eigenschaften einer Druckwarteschlange ohne Reflektion in Druckwarteschlange zur Verwaltung
            PrintQueue druckwarteschlange = LocalPrintServer.GetDefaultPrintQueue();

            // Druckprozess definieren und starten
            Process drucken = new Process();
            drucken.StartInfo.FileName = @"C:\Users\homsi\AppData\Local\SumatraPDF\SumatraPDF.exe";
            drucken.StartInfo.UseShellExecute = true;
            drucken.StartInfo.CreateNoWindow = true;
            drucken.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            drucken.StartInfo.Arguments = "-print-to \"" + pDruckername + "\" -exit-when-done \"" + pPfad + "\"";
            drucken.Start();

            // DruckauftragsInfo
            PrintSystemJobInfo druckauftrag = null;
            // Solange der Druckprozess lauft & DruckauftragsInfo noch nicht definiert & der Prozess nicht abgebrochen
            while(!drucken.HasExited && druckauftrag == null && !pCancellationToken.IsCancellationRequested) 
            {
                druckwarteschlange.Refresh();
                // Infos der Druckauftraege in der Warteschlange
                PrintJobInfoCollection druckauftraege = druckwarteschlange.GetPrintJobInfoCollection();
                // Solange der Druckauftrag noch nicht fertig
                foreach(PrintSystemJobInfo auftrag in druckauftraege)
                {
                    if (auftrag.Name == pPfad)
                        druckauftrag = auftrag;
                }
                Thread.Sleep(50);
            }

            // Solange der noch nicht Auftrag erfolgreich oder abgebrochen
            while (druckauftrag != null && pCancellationToken.IsCancellationRequested)
            {
                // Warten bis der Auftrag erfolgreich wird / abgebrochen
                Thread.Sleep(50);
            }

            // Wenn der Auftrag abgebrochen
            if (pCancellationToken.IsCancellationRequested)
                return false;
            else
                return true;
        }

        private static bool Print_DOCX(string pPfad, string pDruckername, CancellationToken pCancellationToken)
        {
            // DOCX in PDF
            string pfad = Convert_DOCXtoPDF(pPfad);
            bool erg = false;
  
            // Wenn Umwandulung erfolgreich
            if (pfad != "")
            {
                // PDF ausdruecken & loeschen
                erg = Print_PDF(pPfad, pDruckername, pCancellationToken);
                File.Delete(pPfad);
            }

            if (erg)
                return true;
            else
                return false;
        }

        private static string Convert_DOCXtoPDF(string pPfad)
        {
            // hinzuegen Projektverweis Microsoft Word Library
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object fehlt = System.Reflection.Missing.Value;

            FileInfo wordDatei = new FileInfo(pPfad);

            word.Visible = false;
            word.ScreenUpdating = false;

            object dateiname = (object)wordDatei.FullName;

            // Word Dokument fokusieren
            Document doc = word.Documents.Open(ref dateiname, ref fehlt,
                ref fehlt, ref fehlt, ref fehlt, ref fehlt, ref fehlt,
                ref fehlt, ref fehlt, ref fehlt, ref fehlt, ref fehlt,
                ref fehlt, ref fehlt, ref fehlt, ref fehlt);
            doc.Activate();

            // Dateiname und Dateiformat in PDF formatieren
            object ausgabeDateiname = wordDatei.FullName.Replace(".docx", ".pdf");
            object dateiFormat = WdSaveFormat.wdFormatPDF;

            // Datei Umwandeln
            doc.SaveAs2(ref ausgabeDateiname,
                ref dateiFormat, ref fehlt, ref fehlt,
                ref fehlt, ref fehlt, ref fehlt, ref fehlt,
                ref fehlt, ref fehlt, ref fehlt, ref fehlt,
                ref fehlt, ref fehlt, ref fehlt, ref fehlt);

            object speichereAnderungen = WdSaveOptions.wdSaveChanges;
            ((_Document)doc).Close(ref speichereAnderungen, ref fehlt, ref fehlt);
            doc = null;

            ((_Application)word).Quit(ref fehlt, ref fehlt, ref fehlt);
            word = null;

            // Ergebnisse der Umwandelung ermitteln
            if (File.Exists(pPfad.Replace(".docx", ".pdf")))
                return pPfad.Replace(".docx", ".pdf");
            else
                return "";
        }

        private static bool PrintTxt(string pPfad, string pDruckername)
        {
            try
            {
                streamToPrint = new StreamReader(pPfad);
                try
                {
                    printFont = new System.Drawing.Font("Arial", 10);
                    PrintDocument pd = new PrintDocument();
                    pd.PrintPage += new PrintPageEventHandler(Pd_PrintPage);
                    pd.Print();
                }
                finally
                {
                    streamToPrint.Close();
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        private static void Pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            float linesPerPage = 0;
            float yPos = 0;
            int count = 0;
            float leftMargin = ev.MarginBounds.Left;
            float topMargin = ev.MarginBounds.Top;
            string line = null;

            // berechnen Zeilenanzahl des Dokuments
            linesPerPage = ev.MarginBounds.Height /
               printFont.GetHeight(ev.Graphics);

            // ausdrucken alle Zeilen
            while (count < linesPerPage && ((line = streamToPrint.ReadLine()) != null))
            {
                yPos = topMargin + (count * printFont.GetHeight(ev.Graphics));
                ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, yPos, new StringFormat());
                count++;
            }

            // Wenn es noch mehr Zeilen gibt => ausdrucken 
            if (line != null)
                ev.HasMorePages = true;
            else
                ev.HasMorePages = false;
        }
    }
}