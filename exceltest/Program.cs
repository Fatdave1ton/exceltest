using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace exceltest
{
	class Program
	{
		static void Main(string[] args)
		{

			int n = 0;
			do
			{
				openExcel();
				Console.WriteLine(n);
				n++;
				//System.Threading.Thread.Sleep(1000);
				
			} while (n < 20);
			Console.ReadKey();			
		}
		static void openExcel()
		{
			Application excelInstance = new Application();
			Workbook test;
			Workbooks
				tmp;

			tmp = excelInstance.Workbooks;

			test = tmp.Open("f:\\helloworld.xlsm");
			excelInstance.Visible = true;
			excelInstance.Run("Sheet1.test");
			excelInstance.Quit();
			System.Runtime.InteropServices.Marshal.ReleaseComObject(test);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(tmp);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(excelInstance);
		}
	}
}
