using System;
using System.IO;  
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;
using System.Text.RegularExpressions;

//C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /t:exe /out:\Users\ctrueman\Desktop\EXECUTE.exe C:\Users\ctrueman\Desktop\program\new.cs /r:C:\Users\ctrueman\Desktop\program\Microsoft.Office.Interop.Excel.dll /r:C:\Users\ctrueman\Desktop\program\Microsoft.Office.Interop.Access.dll
 
namespace RecordsUpdate
{
	public class Program
	{
		public static void Main()
		{	
			CustomDataArrays customdataarrays = new CustomDataArrays();
			customdataarrays.PopulateData();
			
			Console.WriteLine("Updates finished...");
			Console.ReadLine();		
		}
	}
	

	public static class ExcelDataExtraction
	{
		
		public static Excel.Workbook OpenExcel(Excel.Application excelApp)
		{

			string[] CompleteFile;

			string path = @"W:\IOS\IA\CCU";
			
			if (excelApp != null)
			{			
			
				CompleteFile = GetDirectory(path);
				Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(CompleteFile[0], 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

				return excelWorkbook;
			}
			else
				return null;
		}
		
		
		
		
		public static string[] GetDirectory(string path)
		{

			string ExcelName = GetString();
			var file = Directory.GetFiles(path, ExcelName, SearchOption.AllDirectories);
			
			if (file.Length == 0)
			{
			
				Console.WriteLine("File doesn't exist!");
				Console.ReadLine();
				Environment.Exit(0);
				return null;
			
			}
			else
			{
				return file;
			
			}
	
		}

		
		public static string GetString()
		{	
			Console.Write("Enter spreadsheet name:");
			string FileName = Console.ReadLine();
	
			
			while (!((Regex.IsMatch(FileName, @"[0-9][0-9]\.[0-9][0-9]\.[0-9][0-9]")) || (Regex.IsMatch(FileName, @"[0-9][0-9]\.[0-9]\.[0-9][0-9]")) || (Regex.IsMatch(FileName, @"[0-9]\.[0-9][0-9]\.[0-9][0-9]")) || (Regex.IsMatch(FileName, @"[0-9]\.[0-9]\.[0-9][0-9]"))))
			{
				Console.WriteLine("Please enter a valid move sheet form!");
				Console.WriteLine("Please enter a spreadsheet file to look for:");
				
				FileName = Console.ReadLine();
			}
			
			FileName = FileName + ".xlsx";
						
			return FileName;
		}
	
				
		
		public static string SplitNameFirst(string[] Name)
		{
			
			
			 if(Name.Length == 1)
				 {	 
					 return Name[0];			
				 }
				 else
				 {
					 return Name[0]; 
	 
				 }	
		
		}
		
		
		public static string SplitNameLast(string[] Name)
		{
			
			 if(Name.Length == 1)
				 {	 
					 return Name[0];	
			
				 }
				 else
				 {
					 return Name[1]; 
	 
				 }	
		
		}
		
		
		public static string DeterminePurpose(int i, string temp)
		{
			string MyString = temp;
			
			 if(MyString == "MPOE")
			 {
				return MyString;

			 }
			 
			 else
				return "DESK PHONE";
	
		}


		
		public static string FloorNumber(int i, string temp)
		{
		
			 string MyString = temp; 
			 char secondletter = MyString[1];

			
			 if(MyString == "MPOE")
			 {				
				return MyString;
			 }
				
			 else if(secondletter>='0' && secondletter<='9')
			 {
				return MyString.Substring(0,2);
				 
			 }
			 
			 else
			 {
				return MyString.Substring(0,1);
				
			 }
		}
		
		
		
		public static int FindLastRow(Excel.Worksheet excelWorksheet)
		{	
			 int Row = 9;   //starting point in excel sheet
			 int Count = 0;
			 bool ValidRow = true;
			 
			 while (ValidRow == true)
			 {	   
				Excel.Range Range = (Excel.Range)excelWorksheet.Cells[Row, 6]; 
				//validates using phone number column
		   
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
		public static OleDbConnection ConnectToDatabase()
		{
			string ConnStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=W:\\IOS\\IA\\CCU\\Service Record DB\\Service.mdb;Jet OLEDB:Database Password=job;";
			OleDbConnection MyConn = new OleDbConnection(ConnStr);
			MyConn.Open();
			
			return MyConn;
		}		
	}
	
		
	public class CustomDataArrays
	{
		 private string p_division;
		 private string p_number;
		 private string p_floor;
		 private string p_riser; 
		 private string p_sio;
		 private string p_cubicle; 
		 private string p_first;
		 private string p_last;
		 private string p_purpose;
		 
		 

			
		 public string Purpose
		 {
			get {return p_purpose;}
			set {p_purpose = value;
					
					if(p_purpose == "")
						p_purpose = " ";
				
				}
		 
		 }
		 
		 public string Division
		 {

			get {return p_division;}
			set {p_division = value;
					
					if(p_division == "")
						p_division = " ";
				
				}
		 
		 }
			
		 public string First
		 {

			get {return p_first;}
			set {p_first = value;
					
					if(p_first == "")
						p_first = " ";
				
				}
		 
		 
		 }
			
		 public string Last
		 {

			get {return p_last;}
			set {p_last = value;
					
					if(p_last == "")
						p_last = " ";
				
				}
		 
		 
		 }
		
		 public string Number
		 {

			get {return p_number;}
			set {p_number = value;
					
					if(p_number == "")
						p_number = " ";
				
				}
		 
		 
		 }
		 
		 public string Floor
		 {
			
			get {return p_floor;}
			set {p_floor = value;
					
					if(p_floor == "")
						p_floor = " ";
				
				}
		 
			
		 }
			
		 public string Riser
		 {

			get {return p_riser;}
			set {p_riser = value;
					
					if(p_riser == "")
						p_riser = " ";
				
				}
		 
			
		 }
			
		 public string Sio
		 {

			get {return p_sio;}
			set {p_sio = value;
					
					if(p_sio == "")
						p_sio = " ";
				
				}
		 
	
		 }
			
		 public string Cubicle 
		 {

			get {return p_cubicle;}
			set {p_cubicle = value;
					
					if(p_cubicle == "")
						p_cubicle = " ";
				
				}
		 
		 
		 }
		
		
		 
		 private CustomDataArrays[] point;	
		 
		 public void InitArray(int rowCount)
		 {
			point = new CustomDataArrays[rowCount]; //initialize array
			
			for(int i = 0; i < rowCount; i++)
			{
				point[i] = new CustomDataArrays(); //initialize each element
			}	
		 }
		 
 

		public void PopulateData()
		{	
			int row = 9;
			int i;
			string[] Name = new string[3];
			string date = DateTime.Now.ToString("MM/dd/yyy");
			OleDbConnection MyConn = DataBase.ConnectToDatabase();
			int changedValues = 0;
			OleDbCommand Modify = null;
			Excel.Application excelApp = new Excel.Application();
			Excel.Workbook excelWorkbook = ExcelDataExtraction.OpenExcel(excelApp);
			Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];
			int rowCount = ExcelDataExtraction.FindLastRow(excelWorksheet);
			
			InitArray(rowCount);
			
			for(i = 0; i < rowCount; i++)
			{			 
			
				 string TempName = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 4]).Value2);
				 Name = TempName.Split(' ');
				 
				 string sio = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 11]).Value2);
				 string temp = sio;				 
				 point[i].Sio = sio;
 				 point[i].First = ExcelDataExtraction.SplitNameFirst(Name);
				 point[i].Last = ExcelDataExtraction.SplitNameLast(Name);
				 point[i].Floor = ExcelDataExtraction.FloorNumber(i,temp);
				 point[i].Purpose = ExcelDataExtraction.DeterminePurpose(i,temp);
				 point[i].Division = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 3]).Value2);
				 point[i].Number = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 6]).Value2);
				 point[i].Riser = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 9]).Value2);
				 point[i].Cubicle = Convert.ToString(((Excel.Range)excelWorksheet.Cells[row, 13]).Value2);
				 point[i].Cubicle = point[i].Cubicle.Trim();
				 try
				 {
					 string cmd = "UPDATE [SERVICE MAIN TABLE] SET Divn='" + point[i].Division + "',[First Name]='" + point[i].First + "',Purpose='" + point[i].Purpose + "',[Last Name]='" + point[i].Last + "', FLOOR='" + point[i].Floor + "',RISER='" + point[i].Riser + "',SIO='" + point[i].Sio + "',CUBICLE='" + point[i].Cubicle +  "',PCAUpdtDt='" + date + "',DivnChgDt='" + date + "' WHERE [Phone Nbr]= '" + point[i].Number + "'";			
					 Modify = new OleDbCommand(cmd, MyConn);
					 changedValues = Modify.ExecuteNonQuery();	
				 }
	
				 catch(Exception ex)
			     {
				  Console.WriteLine(ex.Message);	
				  
				 }
				 
			     row++;			
			}	
			
			MyConn.Close();
			excelWorkbook.Close();
			excelApp.Quit();
		}
 
 
	}


	
	
}
	
