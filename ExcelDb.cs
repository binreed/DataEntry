using ClosedXML.Excel;

namespace PulledPropertyApp;

public class ExcelDb
{
    private readonly string _dbPath;
    private const string RecordsSheet = "Records";
    private const string AttachmentsSheet = "Attachments";

    public ExcelDb(string dbPath)
    {
        _dbPath = dbPath;
        EnsureDb();
    }

    private void EnsureDb()
    {
        if (File.Exists(_dbPath)) return;

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add(RecordsSheet);

        var headers = new[]
        {
            "Id","CreatedAt",
            "PulledRefNo","PulledArea","PulledDistrict","PulledPlotNo","PulledUsageType",
            "DecreeNo","DecreeDate","DecreeSource","PrevOwner","CurrOwner",
            "AltRefNo","AltArea","AltDistrict","AltPlotNo","AltUsageType",
            "DecreeStatus"
        };

        for (int i = 0; i < headers.Length; i++)
            ws.Cell(1, i + 1).Value = headers[i];

        ws.Range(1, 1, 1, headers.Length).Style.Font.Bold = true;
        ws.Columns().AdjustToContents();

        var wa = wb.Worksheets.Add(AttachmentsSheet);
        var aHeaders = new[] { "AttachmentId","RecordId","PulledRefNo","FileName","FilePath","AddedAt" };
        for (int i = 0; i < aHeaders.Length; i++)
            wa.Cell(1, i + 1).Value = aHeaders[i];

        wa.Range(1, 1, 1, aHeaders.Length).Style.Font.Bold = true;
        wa.Columns().AdjustToContents();

        wb.SaveAs(_dbPath);
    }

    public RecordModel? FindByPulledRef(string pulledRefNo)
    {
        using var wb = new XLWorkbook(_dbPath);
        var ws = wb.Worksheet(RecordsSheet);

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (int r = 2; r <= lastRow; r++)
        {
            var val = ws.Cell(r, 3).GetString();
            if (string.Equals(val, pulledRefNo, StringComparison.OrdinalIgnoreCase))
                return ReadRecord(ws, r);
        }
        return null;
    }

    public int UpsertByPulledRef(RecordModel model)
    {
        if (string.IsNullOrWhiteSpace(model.PulledRefNo))
            throw new ArgumentException("الرقم المرجعي للعقار المسحوب مطلوب.");

        using var wb = new XLWorkbook(_dbPath);
        var ws = wb.Worksheet(RecordsSheet);

        int targetRow = -1;
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;

        for (int r = 2; r <= lastRow; r++)
        {
            var val = ws.Cell(r, 3).GetString();
            if (string.Equals(val, model.PulledRefNo, StringComparison.OrdinalIgnoreCase))
            {
                targetRow = r;
                break;
            }
        }

        if (targetRow == -1)
        {
            targetRow = lastRow + 1;
            model.Id = GetNextId(ws);
            model.CreatedAt = DateTime.Now;
        }
        else
        {
            model.Id = ws.Cell(targetRow, 1).GetValue<int>();
            model.CreatedAt = ws.Cell(targetRow, 2).GetDateTime();
        }

        WriteRecord(ws, targetRow, model);
        wb.Save();
        return model.Id;
    }

    private int GetNextId(IXLWorksheet ws)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        int maxId = 0;
        for (int r = 2; r <= lastRow; r++)
        {
            if (int.TryParse(ws.Cell(r, 1).GetString(), out var id))
                maxId = Math.Max(maxId, id);
        }
        return maxId + 1;
    }

    private static RecordModel ReadRecord(IXLWorksheet ws, int r)
    {
        DateTime? decreeDate = null;
        if (!ws.Cell(r, 9).IsEmpty())
        {
            if (ws.Cell(r, 9).TryGetValue<DateTime>(out var dt))
                decreeDate = dt;
        }

        return new RecordModel
        {
            Id = ws.Cell(r, 1).GetValue<int>(),
            CreatedAt = ws.Cell(r, 2).GetDateTime(),

            PulledRefNo = ws.Cell(r, 3).GetString(),
            PulledArea = ws.Cell(r, 4).GetString(),
            PulledDistrict = ws.Cell(r, 5).GetString(),
            PulledPlotNo = ws.Cell(r, 6).GetString(),
            PulledUsageType = ws.Cell(r, 7).GetString(),

            DecreeNo = ws.Cell(r, 8).GetString(),
            DecreeDate = decreeDate,
            DecreeSource = ws.Cell(r, 10).GetString(),
            PrevOwner = ws.Cell(r, 11).GetString(),
            CurrOwner = ws.Cell(r, 12).GetString(),

            AltRefNo = ws.Cell(r, 13).GetString(),
            AltArea = ws.Cell(r, 14).GetString(),
            AltDistrict = ws.Cell(r, 15).GetString(),
            AltPlotNo = ws.Cell(r, 16).GetString(),
            AltUsageType = ws.Cell(r, 17).GetString(),

            DecreeStatus = ws.Cell(r, 18).GetString()
        };
    }

    private static void WriteRecord(IXLWorksheet ws, int r, RecordModel m)
    {
        ws.Cell(r, 1).Value = m.Id;
        ws.Cell(r, 2).Value = m.CreatedAt;

        ws.Cell(r, 3).Value = m.PulledRefNo;
        ws.Cell(r, 4).Value = m.PulledArea;
        ws.Cell(r, 5).Value = m.PulledDistrict;
        ws.Cell(r, 6).Value = m.PulledPlotNo;
        ws.Cell(r, 7).Value = m.PulledUsageType;

        ws.Cell(r, 8).Value = m.DecreeNo;
        ws.Cell(r, 9).Value = m.DecreeDate.HasValue ? m.DecreeDate.Value : "";
        ws.Cell(r, 10).Value = m.DecreeSource;
        ws.Cell(r, 11).Value = m.PrevOwner;
        ws.Cell(r, 12).Value = m.CurrOwner;

        ws.Cell(r, 13).Value = m.AltRefNo;
        ws.Cell(r, 14).Value = m.AltArea;
        ws.Cell(r, 15).Value = m.AltDistrict;
        ws.Cell(r, 16).Value = m.AltPlotNo;
        ws.Cell(r, 17).Value = m.AltUsageType;

        ws.Cell(r, 18).Value = m.DecreeStatus;

        ws.Columns().AdjustToContents();
    }
}
