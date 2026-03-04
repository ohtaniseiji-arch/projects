using System;
using System.Runtime.InteropServices;

class Program
{
    static void Main()
    {
        string path = @"C:\test\経営情報シート.xlsm";

        Type? excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType is null)
        {
            Console.WriteLine("Excel が見つかりません。Microsoft Excel がインストールされているか確認してください。");
            return;
        }

        dynamic app = Activator.CreateInstance(excelType)!;
        app.Visible = false;

        dynamic wb = app.Workbooks.Open(path);
        dynamic ws = wb.Sheets["経営情報シート"];
        dynamic range = ws.Range["N5", "DK11"];

        foreach (dynamic cell in range.Cells)
        {
            if (Convert.ToBoolean(cell.HasFormula))
            {
                object formula = cell.Formula;
                object recalculated = ws.Evaluate(formula);
                object value = cell.Value;

                if (!Equals(value, recalculated))
                {
                    Console.WriteLine(
                        $"不一致: {cell.Address}  value={value}  calc={recalculated}");
                }
            }

            Marshal.ReleaseComObject(cell);
        }

        Marshal.ReleaseComObject(range);
        Marshal.ReleaseComObject(ws);
        wb.Close(false);
        Marshal.ReleaseComObject(wb);
        app.Quit();
        Marshal.ReleaseComObject(app);

        Console.WriteLine("検証完了");
    }
}