using System;
using System.Collections.Generic;

namespace VirtualIT.HardwareProbe
{
    public class HardwareProbeSnapshot
    {
        public DateTime TimestampUtc { get; set; } = DateTime.UtcNow;
        public string Status { get; set; } = "OK";
        public MotherboardInfo Motherboard { get; set; } = new();
        public CpuInfo Cpu { get; set; } = new();
        public GpuInfo Gpu { get; set; } = new();
        public List<StorageInfo> Storage { get; set; } = new();
        public List<FanInfo> Fans { get; set; } = new();
        public List<string> Errors { get; set; } = new();
    }

    public class MotherboardInfo
    {
        public string? Vendor { get; set; }
        public string? Model { get; set; }
        public string? Serial { get; set; }
    }

    public class CpuInfo
    {
        public double? TemperatureC { get; set; }
    }

    public class GpuInfo
    {
        public string? Name { get; set; }
        public double? TemperatureC { get; set; }
        public double? LoadPercent { get; set; }
        public int? VramTotalMB { get; set; }
        public int? VramUsedMB { get; set; }
    }

    public class StorageInfo
    {
        public string? Device { get; set; }
        public string? Model { get; set; }
        public double? TemperatureC { get; set; }
    }

    public class FanInfo
    {
        public string? Name { get; set; }
        public int? Rpm { get; set; }
    }
}
