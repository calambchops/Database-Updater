using System;
using System.IO;  
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;
//C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /t:exe /out:\Users\ctrueman\Desktop\EXECUTE.exe C:\Users\ctrueman\Desktop\program\new.cs /r:C:\Users\ctrueman\Desktop\program\Microsoft.Office.Interop.Excel.dll /r:C:\Users\ctrueman\Desktop\program\Microsoft.Office.Interop.Access.dll
 
 
namespace RecordsUpdate
{
	public class Program
	{
	
		public static void Main()
		{
			
			ExcelDataExtraction.OpenExcel();
			
			DataBase.ConnectToDatabase();
			
			Console.ReadLine();
			
		}
	
	}
	

	public class ExcelDataExtraction
	{
	
		public static int rowCount;
				
	
		public static void OpenExcel()
		{
		
			Excel.Application excelApp = new Excel.Application();
		
			if (excelApp != null)
			{			
			
				Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\Users\ctrueman\Desktop\program\test.xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
				Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];
				
		
				rowCount =  FindLastRow(excelWorksheet);
				CustomDataArrays.InitValues(rowCount);				
				SaveData(excelWorksheet, rowCount);
				 		 
		
				excelWorkbook.Close();
				excelApp.Quit();
 
			}
			
									
		}
	
				
		static void SaveData(Excel.Worksheet excelWorksheet, int rowCount)
		{
		
			int row = 9;
			string[] Name = new string[3];
			
			for(int i=0; i < rowCount; i++)
			{
				 
				 string TempName;
				
				 TempName = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 4]).Value2);
				 Name = TempName.Split(' ');

				 CustomDataArrays.first[i] = Name[0];
				 CustomDataArrays.last[i] = Name[1];
				 
				 CustomDataArrays.division[i] = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 3]).Value2);
				 CustomDataArrays.number[i] = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 6]).Value2);
				 CustomDataArrays.riser[i] = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 9]).Value2);
				 CustomDataArrays.sio[i] = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 11]).Value2);
				 CustomDataArrays.cubicle[i] = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 13]).Value2);
				
				 string temp = CustomDataArrays.sio[i] = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 11]).Value2);
				 
				 FloorNumber(i,temp);
				
				 Console.WriteLine(CustomDataArrays.division[i]);
				 Console.WriteLine(CustomDataArrays.number[i]);
				 Console.WriteLine(CustomDataArrays.floor[i]);
				 Console.WriteLine(CustomDataArrays.riser[i]);
				 Console.WriteLine(CustomDataArrays.sio[i]);
				 Console.WriteLine(CustomDataArrays.cubicle[i]);
						
				row++;
				
			}
				
		}
		
		
				
		static void FloorNumber(int i, string temp)
		{
			
			 string MyString = temp;
			
			 if(MyString == "MPOE")
			 {
				
				CustomDataArrays.floor[i] = MyString;
				return;
			 }
				
			 CustomDataArrays.floor[i] = MyString.Substring(0,1);
			

		}
		
		
		
		static int FindLastRow(Excel.Worksheet excelWorksheet)
		{
		
			 int Row = 9;   //starting point in excel sheet
			 int Count = 0;
			 bool ValidRow = true;
			 
			 while (ValidRow == true)
			 {
			   
				Excel.Range Range = (Excel.Range)excelWorksheet.Cells[Row, 6]; 
				//validates using phone number column, 6
		   
				if (Range.Value != null)
				{
				
					Count++;
				}
				else
				{
				
				  ValidRow = false;
				}
				
				Row++;
				
			 }
				
			  return (Count);
			  
		}
		
	}	
	
	
	public class DataBase
	{
	
		public static void ConnectToDatabase()
		{
		
			string ConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\ctrueman\\Desktop\\test.mdb;Jet OLEDB:Database Password=job;";
			OleDbConnection MyConn = new OleDbConnection(ConnStr);
			MyConn.Open();
			
			int changedValues = 0;
			OleDbCommand Modify = null;
			

			try
			{
				string date = DateTime.Now.ToString("MM/dd/yyy");
				
				Console.WriteLine("PRINTING TIME:");
			
				for(int j = 0; j < ExcelDataExtraction.rowCount; j++)
				{
					
					Console.WriteLine("HIBOI");						  
					string cmd = "UPDATE [SERVICE MAIN TABLE] SET Divn='" + CustomDataArrays.division[j] + "',FirstName='" + CustomDataArrays.first[j] + "',LastName='" + CustomDataArrays.last[j] + "', FLOOR='" + CustomDataArrays.floor[j] + "',RISER='" + CustomDataArrays.riser[j] + "',SIO='" + CustomDataArrays.sio[j] + "',CUBICLE='" + CustomDataArrays.cubicle[j] + "',PCAUpdtDt='" + date + "',DivnChgDt='" + date + "' WHERE PhoneNbr= '" + CustomDataArrays.number[j] + "'";			
					Modify = new OleDbCommand(cmd, MyConn);
					changedValues = Modify.ExecuteNonQuery();	
					Console.WriteLine(changedValues);
					
				}
				Console.WriteLine("after console writes:");
				
				
				
				
				
						
			}
			
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			
			}
			
			finally
			{
				MyConn.Close();
			
			}
		
		
		}		
			
			

	}
		
		
	
	
	public class CustomDataArrays
	{
	

		public static int row;
	
	
		 public static void InitValues(int rowCount)
		 {
			
			row = rowCount;
			p_division = new string[row];
		    p_number = new string[row];
		    p_floor = new string[row];
		    p_riser = new string[row]; 
			p_sio = new string[row];
			p_cubicle = new string[row]; 
			p_first = new string[row];
			p_last = new string[row];
		 }
		 
		 public static void print(){
			
			Console.WriteLine(CustomDataArrays.row);
			
		 }
		 
		 
		 private static string[] p_division;
		 private static string[] p_number;
		 private static string[] p_floor;
		 private static string[] p_riser; 
		 private static string[] p_sio;
		 private static string[] p_cubicle; 
		 private static string[] p_first;
		 private static string[] p_last;
		  
		
		 public static string[] division{
		 
			get {return p_division;}
			set {p_division = value;}
		 
			}
			
		 public static string[] first{
		 
			get {return p_first;}
			set {p_first = value;}
		 
			}
			
		 public static string[] last{
		 
			get {return p_last;}
			set {p_last = value;}
		 
			}
		
		 public static string[] number{
		 
			get {return p_number;}
			set {p_number = value;}
		 
			}
		 
		 public static string[] floor{

			get {return p_floor;}
			set {p_floor = value;}
			
			}
			
		 public static string[] riser{
		 
			get {return p_riser;}
			set {p_riser = value;}
			
			}
			
		 public static string[] sio{
		 
			get {return p_sio;}
			set {p_sio = value;}
	
			}
			
		 public static string[] cubicle {
		 
			get {return p_cubicle;}
			set {p_cubicle = value;}
		 
			}
			

		 
	
	}
	
	
	
}
	
