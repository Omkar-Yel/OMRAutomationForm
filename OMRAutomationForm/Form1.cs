using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CorelDRAW;
using VGCore;

namespace OMRAutomationForm
{
    public partial class Form1 : Form
    {
        public static class Prompt
        {
            public static string ShowDialog(string text, string caption)
            {
                Form prompt = new Form()
                {
                    Width = 300,
                    Height = 150,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    Text = caption,
                    StartPosition = FormStartPosition.CenterScreen
                };

                Label textLabel = new Label() { Left = 20, Top = 20, Text = text, AutoSize = true };
                TextBox inputBox = new TextBox() { Left = 20, Top = 50, Width = 240 };

                Button confirmation = new Button() { Text = "OK", Left = 180, Width = 80, Top = 80, DialogResult = DialogResult.OK };
                confirmation.Click += (sender, e) => { prompt.Close(); };

                prompt.Controls.Add(textLabel);
                prompt.Controls.Add(inputBox);
                prompt.Controls.Add(confirmation);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog() == DialogResult.OK ? inputBox.Text : string.Empty;
            }
        }

        private VGCore.Application corelApp;
        private VGCore.Document doc;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        /*private void StartLongTaskButton_Click(object sender, EventArgs e)
        {
            // Show the loading panel
            loadingPanel.Visible = true;

            // Run the task in the background
            Task.Run(() =>
            {
                // Simulate a long-running task
                System.Threading.Thread.Sleep(5000);

                // After the task is done, update the UI on the main thread
                Invoke(new Action(() =>
                {
                    // Hide the loading panel
                    loadingPanel.Visible = false;

                    // Inform the user the task is complete
                    MessageBox.Show("Task Completed!");
                }));
            });
        }*/

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                corelApp = new VGCore.Application();
                corelApp.Visible = true;

                if (corelApp.Documents.Count > 0)
                {
                    doc = corelApp.ActiveDocument;  //Get the Active Document
                }
                else
                {
                    doc = corelApp.CreateDocument();
                    doc.ActivePage.SizeWidth = 210;   // Width in mm for A4
                    doc.ActivePage.SizeHeight = 297;  // Height in mm for A4
                }
                MessageBox.Show("Document Exists");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error : {ex.Message}");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            double left_x_reference = 7.482;
            double left_y_reference = 35.257;
            double circle_x_reference = 21.17;
            double circle_y_reference = 169.773;
            //List<VGCore.Shape> shapes = new List<VGCore.Shape>();
            doc.Unit = VGCore.cdrUnit.cdrMillimeter; // Set units to millimeters
            try
            {
                //Create UDISE Rectangle
                var rect = doc.ActiveLayer.CreateRectangle(16.908, 166.403, 65.648, 221.88);
                rect.Name = "UDISE Box";
                left_x_reference = 19.17;
                left_y_reference = 212.582;

                for (int i = 0; i < 11; i++)
                {
                    rect = doc.ActiveLayer.CreateRectangle(left_x_reference, left_y_reference, left_x_reference + 4, left_y_reference + 5);   //Create UDISE Manual FIlling Boxes
                    rect.Name = "Manual_UDISE_Box_" + i;
                    left_x_reference += 4;
                    if (i == 10)
                    {
                        left_x_reference = 7.482;
                        left_y_reference = 35.257;
                    }
                }

                for (int i = 10; i > -1; i--)
                {
                    //Create UDISE Circle Rows
                    for (int j = 9; j > -1; j--)
                    {
                        //Create UDISE Circle Columns
                        var circle = doc.ActiveLayer.CreateEllipse(circle_x_reference - 1.5, circle_y_reference - 1.5, circle_x_reference + 1.5, circle_y_reference + 1.5);
                        circle.Name = "UDISE_Circle_row_" + i + "_" + j;
                        circle_y_reference += 4.217;
                        //Create numbers for the Circles
                        var text = doc.ActiveLayer.CreateArtisticText(circle_x_reference - 0.635, circle_y_reference - 5.028, j.ToString());
                        text.Text.Story.Font = "Arial";
                        text.Text.Story.Size = 6.5f;
                        text.Name = "UDISE_Column_"+i+"_"+j;
                        text.Fill.UniformColor.RGBAssign(137, 137, 137);
                        if (j == 0)
                        {
                            //Set Circle top reference at the end of the 
                            circle_y_reference = 169.773;
                        }
                    }
                    circle_x_reference += 4.022;

                }

                for (int i = 0; i < 47; i++)
                {
                    //Create Left Timers
                    rect = doc.ActiveLayer.CreateRectangle(left_x_reference, left_y_reference, left_x_reference + 4, left_y_reference + 1.27);
                    rect.Fill.UniformColor.RGBAssign(0, 0, 0);
                    rect.Name = "Left_Timer_No_" + (47 - i);
                    left_y_reference += 4.217;
                    if (i == 0)
                    {
                        //Create bottom end circle timer
                        var circle = doc.ActiveLayer.CreateEllipse((left_x_reference + 2) - 2.397, left_y_reference - (3 * 4.217) - 2.397, (left_x_reference + 2) + 2.397, left_y_reference - (3 * 4.217) + 2.397);
                        circle.Fill.UniformColor.RGBAssign(0, 0, 0);
                        circle.Name = "Bottom_Left_Circle_Timer";
                    }
                    if (i == 46)
                    {
                        //Create top end circle timer
                        var circle = doc.ActiveLayer.CreateEllipse((left_x_reference + 2) - 2.397, left_y_reference + 4.217 - 2.397, (left_x_reference + 2) + 2.397, left_y_reference + 4.217 + 2.397);
                        circle.Fill.UniformColor.RGBAssign(0, 0, 0);
                        //Set reference position for right timers
                        circle.Name = "Top_Left_Circle_Timer";
                        left_y_reference = 35.257;
                        left_x_reference = 198.618;
                    }

                }

                for (int i = 0; i < 47; i++)
                {
                    //Create Right Timers
                    rect = doc.ActiveLayer.CreateRectangle(left_x_reference, left_y_reference, left_x_reference + 4, left_y_reference + 1.27);
                    rect.Fill.UniformColor.RGBAssign(0, 0, 0);
                    left_y_reference += 4.217;
                    if (i == 0)
                    {
                        //Create bottom end circle timer
                        var circle = doc.ActiveLayer.CreateEllipse((left_x_reference + 2) - 2.397, left_y_reference - (3 * 4.217) - 2.397, (left_x_reference + 2) + 2.397, left_y_reference - (3 * 4.217) + 2.397);
                        circle.Fill.UniformColor.RGBAssign(0, 0, 0);
                    }
                    if (i == 46) //Check for last condition
                    {
                        //Create top end circle timer
                        var circle = doc.ActiveLayer.CreateEllipse((left_x_reference + 2) - 2.397, left_y_reference + 4.217 - 2.397, (left_x_reference + 2) + 2.397, left_y_reference + 4.217 + 2.397);
                        circle.Fill.UniformColor.RGBAssign(0, 0, 0);
                        //Reset left and right reference
                        left_y_reference = 219.561;
                        left_x_reference = 16.908;
                    }

                }

                //Create Timer end circles

                //Create UDISE Text Placeholder
                rect = doc.ActiveLayer.CreateRectangle(left_x_reference, left_y_reference, left_x_reference + 48.74, left_y_reference + 4.431);
                rect.Fill.UniformColor.RGBAssign(157, 157, 157);

                var text2 = doc.ActiveLayer.CreateArtisticText(28.226, 220.387, "School UDISE Code");
                text2.Text.Story.Font = "Arial";
                text2.Text.Story.Size = 9f;
                text2.Text.Story.Bold = true;
                // Align the text
                //text2.SetPosition(centerX, centerY);

                MessageBox.Show("UDISE Created");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string user_col_Input = Prompt.ShowDialog("Enter No of columns for question:", "User Input");
            doc.Unit = VGCore.cdrUnit.cdrMillimeter;
            // Handle the input for columns
            int colCount = 0; // Default value in case of invalid input
            if (!string.IsNullOrEmpty(user_col_Input) && int.TryParse(user_col_Input, out colCount))
            {
                MessageBox.Show($"You entered: {colCount} columns");
            }
            else
            {
                MessageBox.Show("Invalid input for columns. Please enter a valid number.");
                return; // Exit the method if input is invalid
            }

            string user_row_Input = Prompt.ShowDialog("Enter No of rows for question:", "User Input");

            // Handle the input for rows
            int rowCount = 0; // Default value in case of invalid input
            if (!string.IsNullOrEmpty(user_row_Input) && int.TryParse(user_row_Input, out rowCount))
            {
                MessageBox.Show($"You entered: {rowCount} rows");
            }
            else
            {
                MessageBox.Show("Invalid input for rows. Please enter a valid number.");
                return; // Exit the method if input is invalid
            }

            double left_x_reference = 16.433;
            double left_y_reference = 35.815;
            double circle_x_reference = 26.043;
            double circle_y_reference = 40.123;
            double adjustment = 0.744;
            double circle_temp_x_reference = circle_x_reference;

            // Now use colCount and rowCount in your loop
            for (int i = 0; i < colCount; i++)
            {
                var rect = doc.ActiveLayer.CreateRectangle(left_x_reference, left_y_reference, left_x_reference + 5.821, left_y_reference + 125.307);
                var rect2 = doc.ActiveLayer.CreateRectangle(left_x_reference + 5.821, left_y_reference, left_x_reference + 5.821 + 23.929, left_y_reference + 125.307);
                circle_x_reference = circle_temp_x_reference;
                for (int j = 0; j < rowCount; j++) 
                {
                    if ((i + 1) * rowCount - j > 10)
                    {
                        adjustment = 0.744 + 0.656;
                    }
                    var text2 = doc.ActiveLayer.CreateArtisticText(circle_x_reference - 6.609 - adjustment, circle_y_reference - 0.821, ((i + 1) * rowCount - j).ToString());
                    text2.Text.Story.Font = "Arial";
                    text2.Text.Story.Size = 6.5f;
                    text2.Text.Story.Bold = true;
                    text2.Text.Story.Fill.UniformColor.RGBAssign(0, 0, 0);
                    for (int k = 1; k < 5; k++)
                    {                        
                        var circle = doc.ActiveLayer.CreateEllipse(circle_x_reference - 1.5, circle_y_reference - 1.5, circle_x_reference + 1.5, circle_y_reference + 1.5);
                        var text = doc.ActiveLayer.CreateArtisticText(circle_x_reference - 0.596, circle_y_reference - 0.821, k.ToString());
                        text.Text.Story.Font = "Arial";
                        text.Text.Story.Size = 6.5f;
                        text.Fill.UniformColor.RGBAssign(137, 137, 137);
                        circle_x_reference += 5.483;
                        if(k == 4)
                        {
                            circle_x_reference = circle_temp_x_reference;
                        }
                    }

                    circle_y_reference += 8.436;
                    if (j == rowCount - 1)
                    {
                        circle_y_reference = 40.123;
                    }
                  
                }
                left_x_reference += 36.85;
                circle_temp_x_reference += 36.851;
            }
            left_x_reference = 120.765;
            left_y_reference = 167.673;
            doc.ActiveLayer.CreateRectangle(left_x_reference, left_y_reference, left_x_reference + 40.183, left_y_reference + 7.984);
            doc.ActiveLayer.CreateRectangle(left_x_reference + 40.183, left_y_reference, left_x_reference + 40.183 + 32.636, left_y_reference + 7.984);
            doc.ActiveLayer.CreateRectangle(left_x_reference, left_y_reference + 7.984, left_x_reference + 40.183 + 32.636, left_y_reference + 7.984 + 41.665 + 7.788);
            MessageBox.Show("Questions Added");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double fixed_Width = 177.151;
            string user_col_Input = Prompt.ShowDialog("Enter No of columns for question:", "User Input");

            // Handle the input for columns
            int colCount = 0; // Default value in case of invalid input
            if (!string.IsNullOrEmpty(user_col_Input) && int.TryParse(user_col_Input, out colCount))
            {
                MessageBox.Show($"You entered: {colCount} columns");
            }
            else
            {
                MessageBox.Show("Invalid input for columns. Please enter a valid number.");
                return; // Exit the method if input is invalid
            }

            string user_row_Input = Prompt.ShowDialog("Enter No of rows for question:", "User Input");

            // Handle the input for rows
            int rowCount = 0; // Default value in case of invalid input
            if (!string.IsNullOrEmpty(user_row_Input) && int.TryParse(user_row_Input, out rowCount))
            {
                MessageBox.Show($"You entered: {rowCount} rows");
            }
            else
            {
                MessageBox.Show("Invalid input for rows. Please enter a valid number.");
                return; // Exit the method if input is invalid
            }

            double left_x_reference = 120.765;
            double left_y_reference = 167.673;
            double circle_x_reference = 26.043;
            double circle_y_reference = 40.123;
            double circle_temp_x_reference = circle_x_reference;

            // Now use colCount and rowCount in your loop
            for (int i = 0; i < colCount; i++)
            {
                circle_x_reference = circle_temp_x_reference;
                for (int j = 0; j < rowCount; j++)
                {
                    for (int k = 1; k < 5; k++)
                    {
                        var circle = doc.ActiveLayer.CreateEllipse(circle_x_reference - 1.5, circle_y_reference - 1.5, circle_x_reference + 1.5, circle_y_reference + 1.5);
                        var text = doc.ActiveLayer.CreateArtisticText(circle_x_reference - 0.596, circle_y_reference - 0.821, k.ToString());
                        text.Text.Story.Font = "Arial";
                        text.Text.Story.Size = 6.5f;
                        text.Fill.UniformColor.RGBAssign(137, 137, 137);
                        circle_x_reference += 5.483;
                        if (k == 4)
                        {
                            circle_x_reference = circle_temp_x_reference;
                        }
                    }

                    circle_y_reference += 8.436;
                    if (j == rowCount - 1)
                    {
                        circle_y_reference = 40.123;
                    }

                }
                circle_temp_x_reference += 36.851;
            }
            MessageBox.Show("Questions Added");
        }
    }
}
