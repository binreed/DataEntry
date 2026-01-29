using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

namespace PulledPropertyApp;

public class MainForm : Form
{
    private readonly ExcelDb _db;
    private readonly string _appDir;
    private readonly string _dbPath;
    private int _currentRecordId = 0;

    // Search
    private TextBox txtSearch = new() { Width = 220 };
    private Button btnSearch = new() { Text = "بحث", Width = 80 };
    private Label lblSearchHint = new() { Text = "بحث بالرقم المرجعي (العقار المسحوب):" };

    // Pulled
    private TextBox pRef = new(); private TextBox pArea = new(); private TextBox pDist = new();
    private TextBox pPlot = new(); private TextBox pUsage = new();

    // Decree
    private TextBox dNo = new(); private DateTimePicker dDate = new() { Format = DateTimePickerFormat.Short, Checked = false, ShowCheckBox = true };
    private TextBox dSource = new(); private TextBox dPrevOwner = new(); private TextBox dCurrOwner = new();

    // Alt
    private TextBox aRef = new(); private TextBox aArea = new(); private TextBox aDist = new();
    private TextBox aPlot = new(); private TextBox aUsage = new();

    // Status
    private ComboBox cmbStatus = new() { DropDownStyle = ComboBoxStyle.DropDownList };

    // Attachments
    private ListBox lstAttach = new() { Height = 120 };
    private Button btnAddAttach = new() { Text = "إضافة مرفق" };
    private Button btnOpenAttach = new() { Text = "فتح" };
    private Button btnRemoveAttach = new() { Text = "حذف من القائمة" };

    // Actions
    private Button btnSave = new() { Text = "حفظ", Width = 100 };
    private Button btnClear = new() { Text = "مسح", Width = 100 };

    public MainForm()
    {
        Text = "نظام إدخال بيانات المرسوم (Excel DB)";
        Width = 980;
        Height = 820;
        StartPosition = FormStartPosition.CenterScreen;
        RightToLeft = RightToLeft.Yes;
        RightToLeftLayout = true;

        _appDir = AppContext.BaseDirectory;
        _dbPath = Path.Combine(_appDir, "Database.xlsx");
        _db = new ExcelDb(_dbPath);

        cmbStatus.Items.AddRange(new[] { "منجز", "غير منجز" });
        cmbStatus.SelectedIndex = 1;

        BuildUi();
        WireEvents();
    }

    private void BuildUi()
    {
        var main = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(12)
        };

        // Search bar
        var searchPanel = new FlowLayoutPanel { Width = 920, AutoSize = true };
        searchPanel.Controls.Add(lblSearchHint);
        searchPanel.Controls.Add(txtSearch);
        searchPanel.Controls.Add(btnSearch);
        main.Controls.Add(searchPanel);

        main.Controls.Add(Section("بيانات العقار المسحوب", new Control[]
        {
            Row("الرقم المرجعي", pRef),
            Row("المنطقة", pArea),
            Row("الحي", pDist),
            Row("رقم القطعة", pPlot),
            Row("نوع الاستخدام", pUsage),
        }));

        main.Controls.Add(Section("بيانات المرسوم", new Control[]
        {
            Row("رقم المرسوم", dNo),
            Row("تاريخ المرسوم", dDate),
            Row("مصدر المرسوم", dSource),
            Row("المالك السابق", dPrevOwner),
            Row("المالك الحالي", dCurrOwner),
        }));

        main.Controls.Add(Section("العقار البديل", new Control[]
        {
            Row("الرقم المرجعي", aRef),
            Row("المنطقة", aArea),
            Row("الحي", aDist),
            Row("رقم القطعة", aPlot),
            Row("نوع الاستخدام", aUsage),
        }));

        main.Controls.Add(Section("المرفقات", new Control[]
        {
            lstAttach,
            new FlowLayoutPanel { AutoSize = true, Controls = { btnAddAttach, btnOpenAttach, btnRemoveAttach } }
        }));

        main.Controls.Add(Section("حالة المرسوم", new Control[]
        {
            Row("الحالة", cmbStatus)
        }));

        var actions = new FlowLayoutPanel { Width = 920, AutoSize = true };
        actions.Controls.Add(btnSave);
        actions.Controls.Add(btnClear);
        main.Controls.Add(actions);

        Controls.Add(main);
    }

    private static GroupBox Section(string title, Control[] content)
    {
        var gb = new GroupBox { Text = title, Width = 920, AutoSize = true, Padding = new Padding(10) };
        var panel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            Dock = DockStyle.Fill
        };

        foreach (var c in content) panel.Controls.Add(c);
        gb.Controls.Add(panel);
        return gb;
    }

    private static Control Row(string label, Control input)
    {
        input.Width = 280;
        var lbl = new Label { Text = label, Width = 160, TextAlign = ContentAlignment.MiddleRight };
        var row = new FlowLayoutPanel { Width = 880, Height = 36, FlowDirection = FlowDirection.RightToLeft };
        row.Controls.Add(input);
        row.Controls.Add(lbl);
        return row;
    }

    private void WireEvents()
    {
        btnSearch.Click += (_, __) => SearchAndLoad();
        btnSave.Click += (_, __) => Save();
        btnClear.Click += (_, __) => ClearForm();

        btnAddAttach.Click += (_, __) => AddAttachment();
        btnOpenAttach.Click += (_, __) => OpenAttachment();
        btnRemoveAttach.Click += (_, __) => RemoveAttachmentFromList();
    }

    private void SearchAndLoad()
    {
        var key = txtSearch.Text.Trim();
        if (string.IsNullOrWhiteSpace(key))
        {
            MessageBox.Show("اكتب الرقم المرجعي للعقار المسحوب للبحث.");
            return;
        }

        var rec = _db.FindByPulledRef(key);
        if (rec == null)
        {
            var res = MessageBox.Show("لا يوجد سجل بهذا الرقم. تريد إنشاء سجل جديد؟", "غير موجود", MessageBoxButtons.YesNo);
            if (res == DialogResult.Yes)
            {
                ClearForm();
                pRef.Text = key;
            }
            return;
        }

        LoadToForm(rec);
        LoadAttachmentsList(rec.PulledRefNo);
    }

    private void LoadToForm(RecordModel r)
    {
        _currentRecordId = r.Id;

        pRef.Text = r.PulledRefNo;
        pArea.Text = r.PulledArea;
        pDist.Text = r.PulledDistrict;
        pPlot.Text = r.PulledPlotNo;
        pUsage.Text = r.PulledUsageType;

        dNo.Text = r.DecreeNo;
        if (r.DecreeDate.HasValue)
        {
            dDate.Value = r.DecreeDate.Value;
            dDate.Checked = true;
        }
        else
        {
            dDate.Checked = false;
        }

        dSource.Text = r.DecreeSource;
        dPrevOwner.Text = r.PrevOwner;
        dCurrOwner.Text = r.CurrOwner;

        aRef.Text = r.AltRefNo;
        aArea.Text = r.AltArea;
        aDist.Text = r.AltDistrict;
        aPlot.Text = r.AltPlotNo;
        aUsage.Text = r.AltUsageType;

        cmbStatus.SelectedItem = (r.DecreeStatus == "منجز") ? "منجز" : "غير منجز";
    }

    private void Save()
    {
        try
        {
            var model = new RecordModel
            {
                PulledRefNo = pRef.Text.Trim(),
                PulledArea = pArea.Text.Trim(),
                PulledDistrict = pDist.Text.Trim(),
                PulledPlotNo = pPlot.Text.Trim(),
                PulledUsageType = pUsage.Text.Trim(),

                DecreeNo = dNo.Text.Trim(),
                DecreeDate = dDate.Checked ? dDate.Value.Date : null,
                DecreeSource = dSource.Text.Trim(),
                PrevOwner = dPrevOwner.Text.Trim(),
                CurrOwner = dCurrOwner.Text.Trim(),

                AltRefNo = aRef.Text.Trim(),
                AltArea = aArea.Text.Trim(),
                AltDistrict = aDist.Text.Trim(),
                AltPlotNo = aPlot.Text.Trim(),
                AltUsageType = aUsage.Text.Trim(),

                DecreeStatus = cmbStatus.SelectedItem?.ToString() ?? "غير منجز"
            };

            var id = _db.UpsertByPulledRef(model);
            _currentRecordId = id;

            // ensure folder exists for attachments
            EnsureAttachmentsFolder(model.PulledRefNo);

            MessageBox.Show("تم الحفظ بنجاح.");
        }
        catch (Exception ex)
        {
            MessageBox.Show("خطأ أثناء الحفظ:\n" + ex.Message);
        }
    }

    private void ClearForm()
    {
        _currentRecordId = 0;
        txtSearch.Clear();

        pRef.Clear(); pArea.Clear(); pDist.Clear(); pPlot.Clear(); pUsage.Clear();
        dNo.Clear(); dDate.Checked = false; dSource.Clear(); dPrevOwner.Clear(); dCurrOwner.Clear();
        aRef.Clear(); aArea.Clear(); aDist.Clear(); aPlot.Clear(); aUsage.Clear();
        cmbStatus.SelectedIndex = 1;

        lstAttach.Items.Clear();
    }

    private string EnsureAttachmentsFolder(string pulledRef)
    {
        var safe = string.IsNullOrWhiteSpace(pulledRef) ? "UNKNOWN" : pulledRef.Trim();
        var dir = Path.Combine(_appDir, "Attachments", safe);
        Directory.CreateDirectory(dir);
        return dir;
    }

    private void LoadAttachmentsList(string pulledRef)
    {
        lstAttach.Items.Clear();
        var dir = EnsureAttachmentsFolder(pulledRef);
        if (!Directory.Exists(dir)) return;

        var files = Directory.GetFiles(dir);
        foreach (var f in files) lstAttach.Items.Add(f);
    }

    private void AddAttachment()
    {
        var pulledRef = pRef.Text.Trim();
        if (string.IsNullOrWhiteSpace(pulledRef))
        {
            MessageBox.Show("لازم تدخل الرقم المرجعي للعقار المسحوب أولاً.");
            return;
        }

        using var ofd = new OpenFileDialog
        {
            Title = "اختر مرفق",
            Multiselect = true
        };

        if (ofd.ShowDialog() != DialogResult.OK) return;

        var dir = EnsureAttachmentsFolder(pulledRef);
        foreach (var src in ofd.FileNames)
        {
            var name = Path.GetFileName(src);
            var dest = Path.Combine(dir, name);

            // avoid overwrite: add (1),(2)...
            dest = MakeUnique(dest);
            File.Copy(src, dest, false);
            lstAttach.Items.Add(dest);
        }
    }

    private static string MakeUnique(string path)
    {
        if (!File.Exists(path)) return path;

        var dir = Path.GetDirectoryName(path)!;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext = Path.GetExtension(path);

        int i = 1;
        while (true)
        {
            var p = Path.Combine(dir, $"{name} ({i}){ext}");
            if (!File.Exists(p)) return p;
            i++;
        }
    }

    private void OpenAttachment()
    {
        if (lstAttach.SelectedItem is not string file || !File.Exists(file))
        {
            MessageBox.Show("اختر مرفق صحيح.");
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = file,
            UseShellExecute = true
        });
    }

    private void RemoveAttachmentFromList()
    {
        if (lstAttach.SelectedIndex >= 0)
            lstAttach.Items.RemoveAt(lstAttach.SelectedIndex);
    }
}

