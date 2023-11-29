using Microsoft.Maui.Controls;
using Microsoft.Maui.Controls.Compatibility;
using System;
using System.Collections.Generic;
using Grid = Microsoft.Maui.Controls.Grid;
using System.IO;
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Reflection;

namespace Laba1_Excel
{

    class Myexception: Exception
    {
        
        public Myexception() { }
        public Myexception(string message) : base(message) { }
    }

    class Calculation
    {
        public string Formula { get; set; }
        public string Result { get; set; }
    }

    public partial class MainPage : ContentPage
    {
         int CountColumn = 20; // кількість стовпчиків (A to Z)
         int CountRow = 50; // кількість рядків


        private List<List<Calculation>> calculations;

        private List<List<Entry>> entryGrid;

        public MainPage()
        {
            InitializeComponent();
            CreateGrid();
            calculations = new List<List<Calculation>>();
            CreateCalculationsGrid();
        }

        //створення таблиці
        private void CreateGrid()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }

        private void CreateCalculationsGrid()
        {
            for (int row = 0; row < CountRow; row++)
            {
                calculations.Add(new List<Calculation>());
            }

            // Add columns 
            for (int col = 0; col < CountColumn; col++)
            {
                for (int row = 0; row < CountRow; row++)
                {
                    calculations[row].Add(new Calculation());
                }
            }
        }

        private void AddColumnsAndColumnLabels()
        {
            // Додати стовпці та підписи для стовпців
            for (int col = 0; col < CountColumn + 1; col++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                if (col > 0)
                {
                    var label = new Label
                    {
                        Text = GetColumnName(col),
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    Grid.SetRow(label, 0);
                    Grid.SetColumn(label, col);
                    grid.Children.Add(label);
                }
            }
            
        }


       
        private void AddRowsAndCellEntries()
        {
            entryGrid = new List<List<Entry>>();

            for (int row = 0; row < CountRow; row++)
            {
                grid.RowDefinitions.Add(new RowDefinition());
                var rowEntries = new List<Entry>();
                // Додати підпис для номера рядка
                var label = new Label
                {
                    Text = (row + 1).ToString(),
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                Grid.SetRow(label, row + 1);
                Grid.SetColumn(label, 0);
                grid.Children.Add(label);

                for (int col = 0; col < CountColumn; col++)
                {
                    var entry = new Entry
                    {
                        Text = "",
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    entry.Unfocused += Entry_Unfocused;

                    Grid.SetRow(entry, row + 1);
                    Grid.SetColumn(entry, col + 1);
                    grid.Children.Add(entry);
                    rowEntries.Add(entry);
                }

                entryGrid.Add(rowEntries);
            }
        }

        


        private string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = string.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;

                            
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        // викликається, коли користувач вийде зі зміненої клітинки(втратить фокус)
    private void Entry_Unfocused(object sender, FocusEventArgs e) {
            var entry = (Entry)sender;
            var row = Grid.GetRow(entry) - 1;
            var col = Grid.GetColumn(entry) - 1;
            var content = entry.Text;
            // Додайте додаткову логіку, яка виконується при виході зі зміненої клітинки
            try
            {
                Check(content);
                entry.Text = content;
                calculations[row][col].Formula = content;
            }
            catch(Myexception ex)
            {
                DisplayAlert("Помилка", "Введено неправильні дані.", "OK");
                entry.Text = "";
            }
            
    }

        private void Check(string formula)
        {
            
            if (formula == "" || formula == null || formula == " ") {
                return;
            }
            formula = formula.Replace(" ", "");
             if(int.TryParse(formula, out int value))
            {
                return;
            }
                           
            else if (formula.StartsWith("="))
            {
                if (IsCellReference(formula.Substring(1)))
                {
                return;
                }
               if(IsValidFormula(formula))
                {
                    return;
                }
                else
                {
                    throw new Myexception("Wrong input");
                }
                 
            }
            else
            {
                
                throw new Myexception("Wrong input"); 
            }
        }

        private bool IsCellReference(string formula)
        {
            return Regex.IsMatch(formula, @"^[A-Z]+:\d+$");
        }

        private bool IsValidFormula(string formula)
        {

            if (!ValidateParentheses(formula)){
                return false;
            }
            if (formula.Contains("mmax") || formula.Contains("mmin"))
            {
                if (!ValidateMinMaxFunction(ref formula))
                {
                    return false;
                }
            }
            if(!ValidateFormula(formula.Substring(1))){
                return false; 
            }
            return true;
        }



        private bool ValidateParentheses(string expression)
        {
            int openParenthesesCount = expression.Count(c => c == '(');
            int closeParenthesesCount = expression.Count(c => c == ')');

            return openParenthesesCount == closeParenthesesCount;
        }

        private bool ValidateFormula(string formula)
        {
            return Regex.IsMatch(formula, @"^(\d+|([A-Z]+:\d+))((\s*[\+\-\*\^/=<>]=?\s*)(\d+|([A-Z]+:\d+)))*$");
        }


        private bool ValidateMinMaxFunction (ref string  formula)
        {
            while (formula.Contains("mmax") || formula.Contains("mmin"))
            {
                int startIndex = -1;
                if (formula.LastIndexOf("mmax") > formula.LastIndexOf("mmin") && formula.LastIndexOf("mmax") != -1 || formula.LastIndexOf("mmin") == -1 && formula.LastIndexOf("mmax") != -1)
                {
                    startIndex = formula.LastIndexOf("mmax");
                }
                else if (formula.LastIndexOf("mmin") > formula.LastIndexOf("mmax") && formula.LastIndexOf("mmin") != -1 || formula.LastIndexOf("mmin") != -1 && formula.LastIndexOf("mmax") == -1)
                {
                    startIndex = formula.LastIndexOf("mmin");
                }

                // Знаходимо початок функції
                int openParenthesisIndex = formula.IndexOf("(", startIndex);

                if (openParenthesisIndex == -1)
                {
                    return false; // Неправильний формат функції
                }

                // Знаходимо кінець функції
                int closeParenthesisIndex = formula.IndexOf(")", openParenthesisIndex);

                if (closeParenthesisIndex == -1)
                {
                    return false; // Неправильний формат функції
                }

                // Видаляємо зайві частини та перевіряємо, чи у дужках є правильний вміст
                string functionContent = formula.Substring(openParenthesisIndex + 1, closeParenthesisIndex - openParenthesisIndex - 1);

                if (!Regex.IsMatch(functionContent, @"^(\d+|([A-Z]+:\d+))((,\s*)?(\d+|([A-Z]+:\d+)))*$"))
                {
                    return false;
                }

                // Видаляємо функцію з виразу
                formula = formula.Remove(startIndex, closeParenthesisIndex - startIndex + 1).Insert(startIndex, "1");
            }
        

            return true;
        }




        private bool result = false;
        private void CalculateButton_Clicked(object sender, EventArgs e)
        {
            Calculate();
            if (!result)
            {
                for (int row = 0; row < CountRow; row++)
                {
                    for (int col = 0; col < CountColumn; col++)
                    {
                        var entry = entryGrid[row][col];
                        entry.Text = calculations[row][col].Result;
                    }
                }// Обробка кнопки "Порахувати"
                result = true;
            }
            else
            {
                for (int row = 0; row < CountRow; row++)
                {
                    for (int col = 0; col < CountColumn; col++)
                    {
                        var entry = entryGrid[row][col];
                        entry.Text = calculations[row][col].Formula;
                    }
                }// Обробка кнопки "Порахувати"
                result = false;
            }
        }



        private void Calculate()
        {
            foreach (var rowCalculations in calculations)
            {
                foreach (var calculation in rowCalculations)
                {
                    // Use your calculation logic here
                    try
                    {
                        calculation.Result = EvaluateFormula(calculation.Formula).ToString();
                    }
                    catch (Myexception ex)
                    {
                        calculation.Result = "Error";
                    }
                }
            }
        }

        private string EvaluateFormula(string formula)
        {
            if(formula == "" || formula == null)
            {
                return "";
            }
            if(int.TryParse(formula, out int value))
            {
                return value.ToString();
            }
            if (formula.StartsWith("="))
            {
                if (IsCellReference(formula.Substring(1)))
                {
                    string cell_result = GetCellValue(formula.Substring(1)) ;
                    if(cell_result != null)
                    {
                        return cell_result;
                    }
                    else
                    {
                        return "0";
                    }
                }
                else
                {
                    string result = CalculateFormula(formula.Substring(1));
                    if(result != null)
                    {
                        return result;
                    }
                    else
                    {
                        return "0";
                    }
                }
               
            }
            else
            {
                throw new Myexception("Error");
            }
        }


        private string GetCellValue(string cellReference)
        {
            var index = cellReference.IndexOf(':');
            var left_side = cellReference.Substring(0, index);
            var right_side = cellReference.Substring(index + 1);
            var new_left_side = left_side.Reverse().ToString();
            int column_number = 0;
            for (int i = 0; i < left_side.Length; i++)
            {
                column_number += i * 26 + RepresentLetterNumber(left_side[i]) ;
            }
            return calculations[int.Parse(right_side)-1][column_number].Result;
        }

        private int RepresentLetterNumber(char letter)
        {
            int aciicode = Convert.ToInt32(letter);
            return aciicode - 65;
        }


        private string CalculateFormula(string formula)
        {
            if (formula.Contains(':')) { 
                   ChangeCellReference(ref formula);
            }
            if(formula.Contains("mmax") || formula.Contains("mmin"))
            {
                OperationsMinAndMax(ref formula);
            }
            if (formula.Contains("^"))
            {
                Pow(ref formula);
            }
            if (formula.Contains("/") || formula.Contains("*"))
            {
                MulDiv(ref formula);
            }
            if(formula.Contains("+")  || formula.Contains("-"))
            {
                PlusMinus(ref formula);
            }
            if (formula.Contains("=") || formula.Contains("<") || formula.Contains(">") || formula.Contains("<=") || formula.Contains(">="))
            {
                Comparison(ref formula);
            }


            return formula;
        }

        private void ChangeCellReference(ref string formula)
        {
            while(formula.Contains(":"))
            {
                int index = formula.IndexOf(":");
                string left_side = "";
                string right_side = "";
                int i = 0;
                int k = 0;
                while (true)
                {
                    if(index - i - 1 < 0)
                    {
                        break;
                    }
                    if (IsLetterAtoZ(formula[index - i -1]))
                    {
                       i++;
                        left_side += formula[index - i];
                         
                    }
                    else
                    {
                       break;

                    }
                }

                while (true)
                {
                    if(index + k + 1 >= formula.Length) { break; }
                    if (IsNumber0to9(formula[index + k+1 ]))
                    {
                        k++;
                        right_side += formula[index + k];
                        
                    }
                    else
                    {
                        break;
                    }
                }

                int column_number = 0;
                for (int j = 0; j < left_side.Length; j++)
                {
                    column_number += j * 26 + RepresentLetterNumber(left_side[j]);
                }
                if(calculations[int.Parse(right_side) - 1][column_number].Result == "true")
                    formula = formula.Remove(index - i, k + i + 1).Insert(index - 1, 1.ToString());
                else if(calculations[int.Parse(right_side) - 1][column_number].Result == "false")
                    formula = formula.Remove(index - i, k + i + 1).Insert(index - 1, 0.ToString());
                else if (calculations[int.Parse(right_side) - 1][column_number].Result == "error")
                    formula = formula.Remove(index - i, k + i + 1).Insert(index - 1, 0.ToString());
                else
                formula = formula.Remove(index - i, k+ i + 1).Insert(index-1 , calculations[int.Parse(right_side) - 1][column_number].Result);
                
            }
        }

        private bool IsLetterAtoZ(char character)
        {
            return char.IsLetter(character) && character >= 'A' && character <= 'Z';
        }

        private bool IsNumber0to9(char character)
        {
            return char.IsDigit(character) && character >= '0' && character <= '9';
        }


        private void OperationsMinAndMax(ref string formula)
        {
            while (formula.Contains("mmax") || formula.Contains("mmin"))
            {
                int startIndex = -1;
                if (formula.LastIndexOf("mmax") > formula.LastIndexOf("mmin") && formula.LastIndexOf("mmax") != -1 || formula.LastIndexOf("mmin") == -1 && formula.LastIndexOf("mmax") != -1)
                {
                    startIndex = formula.LastIndexOf("mmax");
                }
                else if (formula.LastIndexOf("mmin") > formula.LastIndexOf("mmax") && formula.LastIndexOf("mmin") != -1 || formula.LastIndexOf("mmin") != -1 && formula.LastIndexOf("mmax") == -1)
                {
                    startIndex = formula.LastIndexOf("mmin");
                }

                int openParenthesisIndex = formula.IndexOf("(", startIndex);

                
                // Знаходимо кінець функції
                int closeParenthesisIndex = formula.IndexOf(")", openParenthesisIndex);

                string functionContent = formula.Substring(openParenthesisIndex + 1, closeParenthesisIndex - openParenthesisIndex - 1);
                string function = formula.Substring(startIndex, openParenthesisIndex );
                int result = ResultOfMinAndMax(functionContent, function);

                formula = formula.Remove(startIndex, closeParenthesisIndex - startIndex + 1).Insert(startIndex, result.ToString());
            }
        }


        private int ResultOfMinAndMax(string input, string function)
        {
            List<int> values = new List<int>();
            string current_value ="";
            for (int i = 0; i < input.Length; i++)
            {
                if (input[i] != ',') {
                    current_value += input[i];
                }
                else
                {
                    values.Add(Convert.ToInt32(current_value));
                    current_value = "";
                }
                if(i == input.Length - 1)
                {
                    values.Add(Convert.ToInt32(current_value));
                }
            }
            if(function.Contains("mma"))
            {
                return values.Max();
            }
            else
            {
                return values.Min();
            }
        }


        private void MulDiv (ref string formula)
        {
            while(formula.Contains("*") || formula.Contains("/")) {
                int index = -1;
                char operation = ' ';
                if(formula.IndexOf("*") < formula.IndexOf("/") && formula.IndexOf("*") != -1 || formula.IndexOf("*") != -1 && formula.IndexOf("/") == -1)
                {
                    index = formula.IndexOf("*");
                    operation = '*';
                }
                else if(formula.IndexOf("/") < formula.IndexOf("*") && formula.IndexOf("/") != -1 || formula.IndexOf("/") != -1 && formula.IndexOf("*") == -1)
                {
                    index = formula.IndexOf("/");
                    operation = '/';
                }
                string left_side = "";
                string right_side = "";

                int i = 0;
                int k = 0;
                while (true)
                {
                    if(index - i - 1 < 0)
                    {
                        break;
                    }
                    if (IsNumber0to9(formula[index - i - 1]))
                    {
                        i++;
                        left_side = formula[index - i] + left_side;

                    }
                    else
                    {
                        break;

                    }
                }

                while (true)
                {
                    if(index + k + 1 >= formula.Length)
                    {
                        break;
                    }
                    if (IsNumber0to9(formula[index + k + 1]) )
                    {
                        k++;
                        right_side += formula[index + k];

                    }
                    else
                    {
                        break;
                    }
                }
                string result = "";
                if (operation == '*')
                    result = (Convert.ToInt32(left_side) * Convert.ToInt32(right_side)).ToString();
                else
                {
                    if (right_side != "0")
                        result = (Convert.ToInt32(left_side) / Convert.ToInt32(right_side)).ToString();
                    else
                        result = "error";
                }
                formula = formula.Remove(index - i, k + i + 1).Insert(index - i, result.ToString());
            }
        }

        private void Pow(ref string formula)
        {
            while (formula.Contains("^"))
            {
                int index = formula.IndexOf("^");


                string left_side = "";
                string right_side = "";

                int i = 0;
                int k = 0;
                while (true)
                {
                    if (index - i - 1 < 0)
                    {
                        break;
                    }
                    if (IsNumber0to9(formula[index - i - 1]))
                    {
                        i++;
                        left_side = formula[index - i] + left_side;

                    }
                    else
                    {
                        break;

                    }
                }

                while (true)
                {
                    if (index + k + 1 >= formula.Length)
                    {
                        break;
                    }
                    if (IsNumber0to9(formula[index + k + 1]))
                    {
                        k++;
                        right_side += formula[index + k];

                    }
                    else
                    {
                        break;
                    }
                }

                formula = formula.Remove(index - i, k + i + 1).Insert(index - i, Math.Pow(Convert.ToInt32(left_side), Convert.ToInt32(right_side)).ToString());
            }
        }


        private void PlusMinus(ref string formula)
        {
            while(formula.Contains("-") || formula.Contains("+")) {
                int index = -1;
                char operation = ' ';
                if (formula.IndexOf("+") < formula.IndexOf("-") && formula.IndexOf("+") != -1 || formula.IndexOf("+") != -1 && formula.IndexOf("-") == -1)
                {
                    index = formula.IndexOf("+");
                    operation = '+';
                }
                else if (formula.IndexOf("-") < formula.IndexOf("+") && formula.IndexOf("-") != -1 || formula.IndexOf("-") != -1 && formula.IndexOf("+") == -1)
                {
                    index = formula.IndexOf("-");
                    operation = '-';
                }
                string left_side = "";
                string right_side = "";

                int i = 0;
                int k = 0;
                while (true)
                {
                    if (index - i - 1 < 0)
                    {
                        break;
                    }
                    if (IsNumber0to9(formula[index - i - 1]))
                    {
                        i++;
                        left_side = formula[index - i] + left_side;

                    }
                    else
                    {
                        break;

                    }
                }

                while (true)
                {
                    if (index + k + 1 >= formula.Length)
                    {
                        break;
                    }
                    if (IsNumber0to9(formula[index + k + 1]))
                    {
                        k++;
                        right_side += formula[index + k];

                    }
                    else
                    {
                        break;
                    }
                }
                string result = "";
                if (operation == '+')
                    result = (Convert.ToInt32(left_side) + Convert.ToInt32(right_side)).ToString();
                else
                
                    result = (Convert.ToInt32(left_side) - Convert.ToInt32(right_side)).ToString();
                    
                formula = formula.Remove(index - i, k + i + 1).Insert(index - i, result.ToString());
            }



        }
       

        private void Comparison(ref string formula)
        {
            int first_index = formula.IndexOf("=");
            int second_index = formula.IndexOf("<");
            int third_index = formula.IndexOf(">");
            
            if (second_index == -1 &&  third_index == -1)
            {
                Equivalent(ref formula, first_index);
            }
            else if(first_index == -1 && third_index == -1)
            {
                Lesser(ref formula, second_index);
            }
            else if(first_index == -1 && second_index == -1)
            {
                Higher(ref formula, third_index);
            }
            else if (first_index != -1 && second_index != -1)
            {
                LessOrEqv(ref formula, second_index, first_index);
            }
            else
            {
                HighOrEqv(ref formula, third_index, first_index);
            }

        }

        private void Equivalent(ref string formula, int index)
        {
            string left_side = "";
            string right_side = "";
            int i = 0;
            int k = 0;
            while (true)
            {
                if (index - i - 1 < 0)
                {
                    break;
                }
                if (IsNumber0to9(formula[index - i - 1]))
                {
                    i++;
                    left_side = formula[index - i] + left_side;

                }
                else
                {
                    break;

                }
            }

            while (true)
            {
                if (index + k + 1 >= formula.Length)
                {
                    break;
                }
                if (IsNumber0to9(formula[index + k + 1]))
                {
                    k++;
                    right_side += formula[index + k];

                }
                else
                {
                    break;
                }
            }

            string result = "";

            if (Convert.ToInt32(left_side) == Convert.ToInt32(right_side))
            {
                result = "true";
            }
            else
                result = "false";
            formula = formula.Remove(index - i, k + i + 1).Insert(index - i, result.ToString());
        }

        private void Lesser( ref string formula, int index)
        {
            string left_side = "";
            string right_side = "";
            int i = 0;
            int k = 0;
            while (true)
            {
                if (index - i - 1 < 0)
                {
                    break;
                }
                if (IsNumber0to9(formula[index - i - 1]))
                {
                    i++;
                    left_side = formula[index - i] + left_side;

                }
                else
                {
                    break;

                }
            }

            while (true)
            {
                if (index + k + 1 >= formula.Length)
                {
                    break;
                }
                if (IsNumber0to9(formula[index + k + 1]))
                {
                    k++;
                    right_side += formula[index + k];

                }
                else
                {
                    break;
                }
            }

            string result = "";

            if (Convert.ToInt32(left_side) < Convert.ToInt32(right_side))
            {
                result = "true";
            }
            else
                result = "false";
            formula = formula.Remove(index - i, k + i + 1).Insert(index - i, result.ToString());
        }


        private void Higher(ref string formula, int index)
        {
            string left_side = "";
            string right_side = "";
            int i = 0;
            int k = 0;
            while (true)
            {
                if (index - i - 1 < 0)
                {
                    break;
                }
                if (IsNumber0to9(formula[index - i - 1]))
                {
                    i++;
                    left_side = formula[index - i] + left_side;

                }
                else
                {
                    break;

                }
            }

            while (true)
            {
                if (index + k + 1 >= formula.Length)
                {
                    break;
                }
                if (IsNumber0to9(formula[index + k + 1]))
                {
                    k++;
                    right_side += formula[index + k];

                }
                else
                {
                    break;
                }
            }

            string result = "";

            if (Convert.ToInt32(left_side) > Convert.ToInt32(right_side))
            {
                result = "true";
            }
            else
                result = "false";
            formula = formula.Remove(index - i, k + i + 1).Insert(index - i, result.ToString());
        }

        private void LessOrEqv(ref string formula, int first_index, int second_index)
        {
            string left_side = "";
            string right_side = "";
            int i = 0;
            int k = 0;
            while (true)
            {
                if (first_index - i - 1 < 0)
                {
                    break;
                }
                if (IsNumber0to9(formula[first_index - i - 1]))
                {
                    i++;
                    left_side = formula[first_index - i] + left_side;

                }
                else
                {
                    break;

                }
            }

            while (true)
            {
                if (second_index + k + 1 >= formula.Length)
                {
                    break;
                }
                if (IsNumber0to9(formula[second_index + k + 1]))
                {
                    k++;
                    right_side += formula[second_index + k];

                }
                else
                {
                    break;
                }
            }

            string result = "";

            if (Convert.ToInt32(left_side) <= Convert.ToInt32(right_side))
            {
                result = "true";
            }
            else
                result = "false";
            formula = formula.Remove(first_index - i, k + i + 2).Insert(first_index - i, result.ToString());
        }


        private void HighOrEqv(ref string formula, int first_index, int second_index)
        {
            string left_side = "";
            string right_side = "";
            int i = 0;
            int k = 0;
            while (true)
            {
                if (first_index - i - 1 < 0)
                {
                    break;
                }
                if (IsNumber0to9(formula[first_index - i - 1]))
                {
                    i++;
                    left_side = formula[first_index - i] + left_side;

                }
                else
                {
                    break;

                }
            }

            while (true)
            {
                if (second_index + k + 1 >= formula.Length)
                {
                    break;
                }
                if (IsNumber0to9(formula[second_index + k + 1]))
                {
                    k++;
                    right_side += formula[second_index + k];

                }
                else
                {
                    break;
                }
            }

            string result = "";

            if (Convert.ToInt32(left_side) >= Convert.ToInt32(right_side))
            {
                result = "true";
            }
            else
                result = "false";
            formula = formula.Remove(first_index - i, k + i + 2).Insert(first_index - i, result.ToString());
        }



        private void SaveButton_Clicked(object sender, EventArgs e)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Worksheet1");

                    for (int row = 0; row < CountRow; row++)
                    {
                        for (int col = 0; col < CountColumn; col++)
                        {
                            var entry = entryGrid[row][col];
                            worksheet.Cell(row + 1, col + 1).Value = entry.Text;
                        }
                    }

                    string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MyExcelFile.xlsx");
                    workbook.SaveAs(filePath);
                    DisplayAlert("Успішно", "Файл Excel збережено", "OK");
                }
            }
            catch (Exception ex)
            {
                DisplayAlert("Помилка", "Не вдалося зберегти файл Excel: " + ex.Message, "OK");
            }
        }
        private void ReadButton_Clicked(object sender, EventArgs e)
        {

            // Обробка кнопки "Прочитати"

            try
            {
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MyExcelFile.xlsx");

                if (File.Exists(filePath))
                {
                    using (var stream = File.OpenRead(filePath))
                    {
                        DisplayFileContent(stream);
                    }
                }
                else
                {
                    DisplayAlert("Помилка", "Файл не знайдено", "OK");
                }
            }
            catch (Exception ex)
            {
                DisplayAlert("Помилка", "Не вдалося відкрити файл: " + ex.Message, "OK");
            }
        }

        private void DisplayFileContent(Stream stream)
{
    try
    {
        using (var workbook = new XLWorkbook(stream))
        {
            var worksheet = workbook.Worksheet(1); // Припускається, що ваш файл має лише один аркуш

            for (int row = 0; row < CountRow; row++)
            {
                for (int col = 0; col < CountColumn; col++)
                {
                    var entry = entryGrid[row][col];
                    var cellValue = worksheet.Cell(row + 1, col + 1).Value.ToString();
                    entry.Text = cellValue;
                    calculations[row][col].Formula = cellValue;
                }
            }
        }
    }
    catch (Exception ex)
    {
        DisplayAlert("Помилка", "Не вдалося обробити вміст файлу: " + ex.Message, "OK");
    }
}


        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти ? ", "Так", "Ні");
            if (answer)
            {
                System.Environment.Exit(0);
            }
        }

        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка", "Лабораторна робота 1. Студента Чалчинського Ярослава. Варіант 46", "OK");
        }


        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;
                grid.RowDefinitions.RemoveAt(lastRowIndex);
                grid.Children.RemoveAt(lastRowIndex * (CountColumn )); // Remove label
                for (int col = 0; col < CountColumn; col++)
                {

                    grid.Children.RemoveAt( lastRowIndex*(CountColumn+1) + col + 1); // Remove entry
                }
                RemoveRow();
            }
        }

        private void RemoveRow()
        {
            calculations.RemoveAt(calculations.Count - 1);
        }


        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int lastColumnIndex = CountColumn - 1;

                
               // grid.Children.RemoveAt(lastColumnIndex * (CountRow)); // Remove label
                                                                           grid.Children.RemoveAt(grid.Children.Count - 1);

               grid.ColumnDefinitions.RemoveAt(lastColumnIndex);

                for (int row = 0; row < CountRow; row++)
                {

                   // var entry = entryGrid[row][lastColumnIndex];
                    
                    entryGrid[row].RemoveAt(lastColumnIndex);
                    //grid.Children.RemoveAt( row * (CountRow +1 ) + (lastColumnIndex +1)); // Remove entry

                }
                
                RemoveColumn();
                CountColumn--;
            }
        }


        private void RemoveColumn()
        {
            foreach (var row in calculations)
            {
                row.RemoveAt(row.Count - 1);
            }
        }


        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            int newRow = grid.RowDefinitions.Count + 1;
            var rowEntries = new List<Entry>();
            // Add a new row definition
            grid.RowDefinitions.Add(new RowDefinition());
            // Add label for the row number
            var label = new Label
            {
                Text = newRow.ToString(),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, newRow);
            Grid.SetColumn(label, 0);
            grid.Children.Add(label);
            // Add entry cells for the new row
            for (int col = 0; col < CountColumn; col++)
            {
                var entry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, newRow);
                Grid.SetColumn(entry, col + 1);
                grid.Children.Add(entry);
                rowEntries.Add(entry);
            }
            entryGrid.Add(rowEntries);
            AddRow();
            CountRow++;
        }

        private void AddRow()
        {
            var newRow = new List<Calculation>();

            for (int col = 0; col < CountColumn; col++)
            {
                newRow.Add(new Calculation());
            }

            calculations.Add(newRow);
        }


        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            int newColumn = grid.ColumnDefinitions.Count ;
            // Add a new column definition
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            // Add label for the column name
            var label = new Label
            {
                Text = GetColumnName(newColumn),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center
            };
            Grid.SetRow(label, 0);
            Grid.SetColumn(label, newColumn);
            grid.Children.Add(label);
            // Add entry cells for the new column
            for (int row = 0; row < CountRow; row++)
            {
                var entry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, row + 1);
                Grid.SetColumn(entry, newColumn);
                grid.Children.Add(entry);
                entryGrid[row].Add(entry);
            }
            CountColumn++;
            AddColumn();
            
        }

        private void AddColumn()
        {
            for (int row = 0; row < CountRow; row++)
            {
                calculations[row].Add(new Calculation());
               // calculations[row].Insert(calculations[row].Count, new Calculation());
            }
        }
    }
}
