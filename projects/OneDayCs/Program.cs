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

        app.CalculateFull();

        dynamic range = ws.Range["N5", "DK11"];

        bool found = false;

        foreach (dynamic cell in range.Cells)
        {
            if (Convert.ToBoolean(cell.HasFormula))
            {
                object value = cell.Value;

                cell.Calculate();

                object recalculated = cell.Value;

                if (!Equals(value, recalculated))
                {
                    Console.WriteLine(
                        $"不一致: {cell.Address} value={value} calc={recalculated}");
                    found = true;
                }
            }

            Marshal.ReleaseComObject(cell);
        }

        if (!found)
        {
            Console.WriteLine("不一致は検出されませんでした");
        }

        Marshal.ReleaseComObject(range);
        Marshal.ReleaseComObject(ws);
        wb.Close(false);
        Marshal.ReleaseComObject(wb);
        app.Quit();
        Marshal.ReleaseComObject(app);
    }
}