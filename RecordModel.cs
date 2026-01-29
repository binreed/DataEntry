using System;

namespace PulledPropertyApp;

public class RecordModel
{
    public int Id { get; set; }
    public DateTime CreatedAt { get; set; }

    // بيانات العقار المسحوب
    public string PulledRefNo { get; set; } = "";
    public string PulledArea { get; set; } = "";
    public string PulledDistrict { get; set; } = "";
    public string PulledPlotNo { get; set; } = "";
    public string PulledUsageType { get; set; } = "";

    // بيانات المرسوم
    public string DecreeNo { get; set; } = "";
    public DateTime? DecreeDate { get; set; }
    public string DecreeSource { get; set; } = "";
    public string PrevOwner { get; set; } = "";
    public string CurrOwner { get; set; } = "";

    // العقار البديل
    public string AltRefNo { get; set; } = "";
    public string AltArea { get; set; } = "";
    public string AltDistrict { get; set; } = "";
    public string AltPlotNo { get; set; } = "";
    public string AltUsageType { get; set; } = "";

    // حالة المرسوم
    public string DecreeStatus { get; set; } = "غير منجز";
}

