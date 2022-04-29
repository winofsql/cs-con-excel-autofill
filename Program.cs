using System;

namespace excel_2020_11_26
{
	class Program
	{
		static void Main(string[] args)
		{
			// Excel アプリケーション
			dynamic ExcelApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
			// Excel のパス
			string path = Environment.CurrentDirectory + @"\sample.xlsx";

			// Excel を表示( 完成したらコメント化 )
			ExcelApp.Visible = true;
			// 警告を出さない
			ExcelApp.DisplayAlerts = false;

			try
			{
				// ****************************
				// ブック追加
				// ****************************
				dynamic Book = ExcelApp.Workbooks.Add();

				// 通常一つのシートが作成されています
				dynamic Sheet = Book.Worksheets(1);

				// ****************************
				// シート名変更
				// ****************************
				Sheet.Name = "C#の処理";

				// ****************************
				// セルに値を直接セット
				// ****************************
				for (int i = 1; i <= 10; i++)
				{
					Sheet.Cells(i, 1).Value = "処理 : " + i;
				}

				// ****************************
				// 1つのセルから
				// AutoFill で値をセット
				// ****************************
				Sheet.Cells(1,2).Value = "子";
				// 基となるセル範囲
				dynamic SourceRange = Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(1, 2));
				// オートフィルの範囲(基となるセル範囲を含む )
				dynamic FillRange = Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(10, 2));
				SourceRange.AutoFill(FillRange);


				// ****************************
				// 保存
				// ****************************
				Book.SaveAs(path);
			}
			catch (Exception ex)
			{
				ExcelApp.Quit();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
				Console.WriteLine(ex.Message);
				return;
			}

			ExcelApp.Quit();
			// 解放
			System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);

			Console.WriteLine("処理を終了します");

		}
	}
}
