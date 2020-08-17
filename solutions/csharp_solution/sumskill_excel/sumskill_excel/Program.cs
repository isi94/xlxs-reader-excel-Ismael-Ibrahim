/*
 * Author: Ismael Ibrahim
 * Date: 17.08.2020
 * Content: Excel xlsx Parser, calculates the sum up the skill value from
 * a given date
 */
using System;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace sumskill_excel
{
    class Program
    {
        static void Main()
        {
            string fileName = "../../Site_Capacity.xlsx";
            FileStream fs;
            fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(0);


            if (sheet != null)
            {

               string  quit_or_next = "";

                while (!string.Equals(quit_or_next, "q"))
                {
                    //Console.Clear();
                    int year_input = 0;
                    int month_input = 0;
                    string skill_input = "";

                    Console.Write("Skill: ");
                    skill_input = Console.ReadLine();

                    bool invalid_date = true;
                    while (invalid_date)
                    {
                        try
                        {
                            Console.Write("Year: ");
                            year_input = Convert.ToInt32(Console.ReadLine());
                            Console.Write("Month: ");
                            month_input = Convert.ToInt32(Console.ReadLine());
                            invalid_date = false;
                        }
                        catch
                        {
                            Console.WriteLine("Date not valid");
                        }
                    }

                    int col_index
                        = getColIndex(sheet, year_input, month_input);
                    if (col_index == 0)
                    {
                        Console.WriteLine("Date not found");
                    }
                    else
                    {
                        float sum = getSum(sheet, col_index, skill_input);
                        Console.WriteLine("Skill result: " + sum);
                    }

                   

                    Console.Write("Enter \"q\" to quit or any to repeat: \n");
                    quit_or_next = Console.ReadLine().Trim();
                }
               
            }

            fs.Close();
        }

    /* sum up the skill from the given column(date) if the skill attribute matches
    returns the sum of the skill*/
        private static float getSum(ISheet sheet, int col_index, string skill_input)
        {
            float sum_skill = 0.0f;
            int n_row = sheet.LastRowNum;

            for (int i = 1; i <= n_row; i++)
            {
                IRow curRow = sheet.GetRow(i);
                if (curRow.GetCell(1).StringCellValue.Trim() == skill_input.Trim())
                {
                    ICell current_cell = curRow.GetCell(col_index);
                    if(current_cell != null)
                    {
                        if (curRow.GetCell(col_index).CellType == CellType.Numeric);
                        {
                            sum_skill = sum_skill 
                                + float.Parse(curRow.GetCell(col_index).ToString());
                        }
                    }  
                }
            }
            return sum_skill;
        }

        /* compares the title row with the given param year and month
        returns the column index */
        private static int getColIndex(ISheet sheet, int year_input,
            int month_input)
        {
            // If first row is table head, i starts from 1
            IRow header_row = sheet.GetRow(0);
            foreach (ICell cell in header_row)
            {
                if (cell.CellType == CellType.Numeric)
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        DateTime date = cell.DateCellValue.Date;
                        int year = date.Year;
                        int month = date.Month;
                        if (year_input == year && month_input == month)
                            return cell.ColumnIndex;
                    }         
            }
            return 0;
        }
    }
}
