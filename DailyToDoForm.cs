using System;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using Xceed.Words.NET;   // DocX ke liye
using Xceed.Document.NET; // TableDesign ke liye
using Font = System.Drawing.Font;
using System.Drawing.Printing;

public class DailyToDoForm : Form
{
    TextBox[] goalBoxes = new TextBox[5];
    DataGridView taskGrid;
    TextBox[] noteBoxes = new TextBox[3];
    Label quoteLabel;
    string saveFolder;

    public DailyToDoForm()
    {
        this.Text = "Daily To Do";
        this.Size = new Size(800, 700);

        // 🔹 App Icon set karo
        this.Icon = new Icon("icon.ico");

        Label heading = new Label
        {
            Text = "Daily To Do",
            Font = new Font("Arial", 18, FontStyle.Bold),
            Location = new Point(20, 20),
            AutoSize = true
        };
        this.Controls.Add(heading);

        Label dateLabel = new Label
        {
            Text = DateTime.Now.ToLongDateString(),
            Location = new Point(600, 25),
            AutoSize = true
        };
        this.Controls.Add(dateLabel);

        // Goals section
        Label goalsLabel = new Label
        {
            Text = "Goals",
            Location = new Point(20, 70),
            AutoSize = true
        };
        this.Controls.Add(goalsLabel);

        for (int i = 0; i < 5; i++)
        {
            goalBoxes[i] = new TextBox
            {
                Location = new Point(20, 100 + i * 30),
                Size = new Size(400, 25)
            };
            this.Controls.Add(goalBoxes[i]);
        }

        // Task Table
        Label taskLabel = new Label
        {
            Text = "To Do Tasks",
            Location = new Point(20, 260),
            AutoSize = true
        };
        this.Controls.Add(taskLabel);

        taskGrid = new DataGridView
        {
            Location = new Point(20, 290),
            Size = new Size(750, 200),
            ColumnCount = 5
        };
        taskGrid.Columns[0].Name = "Task";
        taskGrid.Columns[1].Name = "Completed";
        taskGrid.Columns[2].Name = "X/-";
        taskGrid.Columns[3].Name = "Cause";
        taskGrid.Columns[4].Name = "Solution";
        taskGrid.AllowUserToAddRows = true;
        this.Controls.Add(taskGrid);

        // Notes
        Label notesLabel = new Label
        {
            Text = "Notes",
            Location = new Point(20, 500),
            AutoSize = true
        };
        this.Controls.Add(notesLabel);

        for (int i = 0; i < 3; i++)
        {
            noteBoxes[i] = new TextBox
            {
                Location = new Point(20, 530 + i * 30),
                Size = new Size(400, 25)
            };
            this.Controls.Add(noteBoxes[i]);
        }

        // Motivational Quote
        quoteLabel = new Label
        {
            Text = GetRandomQuote(),
            Font = new Font("Arial", 10, FontStyle.Italic),
            Location = new Point(20, 630),
            AutoSize = true
        };
        this.Controls.Add(quoteLabel);

        // Buttons
        Button saveButton = new Button
        {
            Text = "Save",
            Location = new Point(450, 530)
        };
        saveButton.Click += SaveToFile;
        this.Controls.Add(saveButton);

        Button pdfButton = new Button
        {
            Text = "Save as PDF",
            Location = new Point(450, 570)
        };
        pdfButton.Click += SaveAsPdf;
        this.Controls.Add(pdfButton);

        // Set default folder
        saveFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "DailyToDo");
        if (!Directory.Exists(saveFolder))
            Directory.CreateDirectory(saveFolder);
    }

    private string GetRandomQuote()
    {
        string[] quotes = {
            "Believe in yourself!",
            "Small steps every day!",
            "You are capable of amazing things!",
            "Stay focused and never give up!",
            "Make today count!"
        };
        Random rnd = new Random();
        return quotes[rnd.Next(quotes.Length)];
    }

    private void SaveToFile(object sender, EventArgs e)
    {
        // 🔹 Save as TXT
        string txtFilePath = Path.Combine(saveFolder, $"ToDo_{DateTime.Now:yyyyMMdd}.txt");
        using (StreamWriter sw = new StreamWriter(txtFilePath))
        {
            sw.WriteLine("Daily To Do - " + DateTime.Now.ToLongDateString());
            sw.WriteLine("\nGoals:");
            foreach (var goal in goalBoxes)
                sw.WriteLine("- " + goal.Text);

            sw.WriteLine("\nTo Do Tasks:");
            foreach (DataGridViewRow row in taskGrid.Rows)
            {
                if (!row.IsNewRow)
                    sw.WriteLine($"{row.Cells[0].Value}, {row.Cells[1].Value}, {row.Cells[2].Value}, {row.Cells[3].Value}, {row.Cells[4].Value}");
            }

            sw.WriteLine("\nNotes:");
            foreach (var note in noteBoxes)
                sw.WriteLine("- " + note.Text);

            sw.WriteLine("\nQuote of the day: " + quoteLabel.Text);
        }

        // 🔹 Save as Word DOCX
        string docxFilePath = Path.Combine(saveFolder, $"ToDo_{DateTime.Now:yyyyMMdd}.docx");
        using (var doc = DocX.Create(docxFilePath))
        {
            doc.InsertParagraph("Daily To Do - " + DateTime.Now.ToLongDateString())
                .FontSize(16).Bold().SpacingAfter(20);

            // Goals
            doc.InsertParagraph("Goals:").Bold();
            foreach (var goal in goalBoxes)
                doc.InsertParagraph("- " + goal.Text);
            doc.InsertParagraph();

            // Tasks as table
            var dataRows = taskGrid.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).ToList();
            if (dataRows.Count > 0)
            {
                var table = doc.AddTable(dataRows.Count + 1, taskGrid.Columns.Count);
                table.Design = TableDesign.ColorfulList;

                // headers
                for (int c = 0; c < taskGrid.Columns.Count; c++)
                    table.Rows[0].Cells[c].Paragraphs[0].Append(taskGrid.Columns[c].Name).Bold();

                // rows
                for (int r = 0; r < dataRows.Count; r++)
                {
                    for (int c = 0; c < taskGrid.Columns.Count; c++)
                    {
                        var val = dataRows[r].Cells[c].Value?.ToString() ?? "";
                        table.Rows[r + 1].Cells[c].Paragraphs[0].Append(val);
                    }
                }

                doc.InsertTable(table);
            }
            doc.InsertParagraph();

            // Notes
            doc.InsertParagraph("Notes:").Bold();
            foreach (var note in noteBoxes)
                doc.InsertParagraph("- " + note.Text);
            doc.InsertParagraph();

            // Quote
            doc.InsertParagraph("Quote of the day: " + quoteLabel.Text).Italic();

            doc.Save();
        }

        MessageBox.Show("Saved:\n" + txtFilePath + "\n" + docxFilePath);
    }

    private void SaveAsPdf(object sender, EventArgs e)
    {
        string pdfFilePath = Path.Combine(saveFolder, $"ToDo_{DateTime.Now:yyyyMMdd}.pdf");

        PrintDocument pd = new PrintDocument();
        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
        pd.PrinterSettings.PrintToFile = true;
        pd.PrinterSettings.PrintFileName = pdfFilePath;

        pd.PrintPage += (s, ev) =>
        {
            float y = 100;
            Font font = new Font("Arial", 10);

            ev.Graphics.DrawString("Daily To Do - " + DateTime.Now.ToLongDateString(),
                new Font("Arial", 14, FontStyle.Bold), Brushes.Black, 100, y);
            y += 40;

            ev.Graphics.DrawString("Goals:", font, Brushes.Black, 100, y);
            y += 20;
            foreach (var goal in goalBoxes)
            {
                ev.Graphics.DrawString("- " + goal.Text, font, Brushes.Black, 120, y);
                y += 20;
            }
            y += 20;

            ev.Graphics.DrawString("To Do Tasks:", font, Brushes.Black, 100, y);
            y += 20;
            foreach (DataGridViewRow row in taskGrid.Rows)
            {
                if (!row.IsNewRow)
                {
                    string line = $"{row.Cells[0].Value}, {row.Cells[1].Value}, {row.Cells[2].Value}, {row.Cells[3].Value}, {row.Cells[4].Value}";
                    ev.Graphics.DrawString(line, font, Brushes.Black, 120, y);
                    y += 20;
                }
            }
            y += 20;

            ev.Graphics.DrawString("Notes:", font, Brushes.Black, 100, y);
            y += 20;
            foreach (var note in noteBoxes)
            {
                ev.Graphics.DrawString("- " + note.Text, font, Brushes.Black, 120, y);
                y += 20;
            }
            y += 20;

            ev.Graphics.DrawString("Quote of the day: " + quoteLabel.Text,
                new Font("Arial", 10, FontStyle.Italic), Brushes.Black, 100, y);
        };

        pd.Print();
        MessageBox.Show("Saved PDF:\n" + pdfFilePath);
    }

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new DailyToDoForm());
    }
}
