using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Xml;
using System.Xml.Linq;

namespace Excel_to_XML_Converter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        //Declare needed strings...
        string SourcePatch;
        string DestinationPatch;
        
        // Create index couter...
        int Index = 0;

        // Create content list...
        List<string> ContentList = new List<string>();
        List<string> HeadersList = new List<string>();

        //Choose the Ecel file to convert.
        private void BrowseSourceButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel CSV Files|*.csv";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                SourceTextBox.Text = openFileDialog.FileName;
                SourcePatch = openFileDialog.FileName;
            }
        }

        //Choose were to save the XAML file.
        private void BrowseDestinationButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML Files|*.xml";
            saveFileDialog.DefaultExt = ".xml";
            if (saveFileDialog.ShowDialog() == true)
            {
                DestinationTextBox.Text = saveFileDialog.FileName;
                DestinationPatch = saveFileDialog.FileName;
            }
        }

        //Convert dokument
        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            //Indlæser og starter XML writeren.
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            XmlWriter writer = XmlWriter.Create(DestinationPatch, settings);

            try
            {
                // Creates the XML document and root-elements...
                writer.WriteStartDocument();
                writer.WriteStartElement("BULK");

                // Use StreamReader to read the lines in the CSV file...
                using (var StreamReader = new StreamReader(SourcePatch))
                {
                    // Read first line and save to HeadersList...
                    var ReadHeader = StreamReader.ReadLine();
                    HeadersList = ReadHeader.Split(';').ToList();

                    // Read content lines in the file...
                    while (!StreamReader.EndOfStream)
                    {
                        // Reset index to 0 for each content line...
                        Index = 0;

                        // Split reading line and save to content list...
                        var Line = StreamReader.ReadLine();
                        ContentList = Line.Split(';').ToList();

                        // Create signature element...
                        writer.WriteStartElement("REGISTRER");
                        writer.WriteStartElement("MEDARBEJDER-SIGNATUR");

                        // Create dynamic user elements...
                        foreach (string Content in ContentList)
                        {
                            // Make ekstra element for "Adresse-linje"... 
                            if (HeadersList.ElementAt(Index) == "ADRESSE-LINJE")
                            {
                                writer.WriteStartElement("ADRESSE-LINJER");
                                writer.WriteElementString(HeadersList[Index], Content);
                                writer.WriteEndElement();
                            }

                            // Create element string for all other dynamic content...
                            else
                            {
                                writer.WriteElementString(HeadersList[Index], Content);
                            }

                            // Increment index by 1 for each index in the reading line...
                            Index++;
                        }

                        // Create static signature type elements and artribute...
                        writer.WriteStartElement("SIGNATUR-TYPE");
                        writer.WriteStartElement("NOEGLEFIL-STRAKSUDSTEDT");
                        writer.WriteStartElement("EMAIL-I-CERTIFIKAT");
                        writer.WriteAttributeString("LDAP-REGISTRERING", "JA");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteEndElement();

                        // Close signature element...
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }

                    // Close the StreamReader...
                    StreamReader.Close();
                }

                // Closing root elements..
                writer.WriteEndElement();

                // End XML Document...
                writer.WriteEndDocument();
                writer.Flush();
                writer.Close();

                //Show succeded messagebox and clear all string and textboxes....
                MessageBox.Show("Konverteringen blev gennemført", "Færdig");
                SourcePatch = null;
                DestinationPatch = null;
                SourceTextBox.Text = null;
                DestinationTextBox.Text = null;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Der skete en fejl");
            }
        }
    }
}
