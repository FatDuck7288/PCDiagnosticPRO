using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using LibreHardwareMonitor.Hardware;

namespace VirtualIT.HardwareProbe
{
    public class HardwareProbeService
    {
        public HardwareProbeSnapshot CollectSnapshot()
        {
            var snapshot = new HardwareProbeSnapshot
            {
                TimestampUtc = DateTime.UtcNow,
                Status = "OK"
            };

            try
            {
                using var computer = new Computer
                {
                    IsCpuEnabled = true,
                    IsGpuEnabled = true,
                    IsMotherboardEnabled = true,
                    IsStorageEnabled = true,
                    IsControllerEnabled = true
                };

                computer.Open();
                var visitor = new UpdateVisitor();
                computer.Accept(visitor);

                PopulateMotherboard(snapshot, computer);
                PopulateCpu(snapshot, computer);
                PopulateGpu(snapshot, computer);
                PopulateStorage(snapshot, computer);
                PopulateFans(snapshot, computer);

                if (snapshot.Errors.Count > 0)
                {
                    snapshot.Status = "WARN";
                }
            }
            catch (Exception ex)
            {
                snapshot.Status = "ERROR";
                snapshot.Errors.Add($"Hardware probe failure: {ex.Message}");
            }

            return snapshot;
        }

        private static void PopulateMotherboard(HardwareProbeSnapshot snapshot, Computer computer)
        {
            var board = QueryBaseBoard();
            if (board != null)
            {
                snapshot.Motherboard.Vendor = board.Value.Vendor;
                snapshot.Motherboard.Model = board.Value.Model;
                snapshot.Motherboard.Serial = board.Value.Serial;
                return;
            }

            var motherboard = computer.Hardware.FirstOrDefault(h => h.HardwareType == HardwareType.Motherboard);
            if (motherboard != null)
            {
                snapshot.Motherboard.Model = motherboard.Name;
            }

            if (string.IsNullOrWhiteSpace(snapshot.Motherboard.Model) &&
                string.IsNullOrWhiteSpace(snapshot.Motherboard.Vendor))
            {
                snapshot.Errors.Add("Motherboard information not available.");
            }
        }

        private static void PopulateCpu(HardwareProbeSnapshot snapshot, Computer computer)
        {
            var cpu = computer.Hardware.FirstOrDefault(h => h.HardwareType == HardwareType.Cpu);
            snapshot.Cpu.TemperatureC = ReadTemperature(cpu, new[] { "Package", "CPU Package", "Core" });

            if (!snapshot.Cpu.TemperatureC.HasValue)
            {
                snapshot.Errors.Add("CPU temperature not available.");
            }
        }

        private static void PopulateGpu(HardwareProbeSnapshot snapshot, Computer computer)
        {
            var gpu = computer.Hardware.FirstOrDefault(h =>
                h.HardwareType == HardwareType.GpuAmd ||
                h.HardwareType == HardwareType.GpuNvidia ||
                h.HardwareType == HardwareType.GpuIntel);

            if (gpu == null)
            {
                snapshot.Errors.Add("GPU not detected.");
                return;
            }

            snapshot.Gpu.Name = gpu.Name;
            snapshot.Gpu.TemperatureC = ReadTemperature(gpu, new[] { "GPU Core", "Core", "Hot Spot", "Hotspot" });
            snapshot.Gpu.LoadPercent = ReadSensorValue(gpu, SensorType.Load, new[] { "GPU Core", "GPU Total", "Core" });

            snapshot.Gpu.VramTotalMB = ReadMemorySensorMb(gpu, new[] { "Memory Total", "GPU Memory Total", "VRAM Total" });
            snapshot.Gpu.VramUsedMB = ReadMemorySensorMb(gpu, new[] { "Memory Used", "GPU Memory Used", "VRAM Used" });

            if (!snapshot.Gpu.TemperatureC.HasValue)
            {
                snapshot.Errors.Add("GPU temperature not available.");
            }

            if (!snapshot.Gpu.LoadPercent.HasValue)
            {
                snapshot.Errors.Add("GPU load not available.");
            }

            if (!snapshot.Gpu.VramTotalMB.HasValue || !snapshot.Gpu.VramUsedMB.HasValue)
            {
                snapshot.Errors.Add("GPU VRAM usage not available.");
            }
        }

        private static void PopulateStorage(HardwareProbeSnapshot snapshot, Computer computer)
        {
            var storageHardware = computer.Hardware.Where(h => h.HardwareType == HardwareType.Storage).ToList();
            if (storageHardware.Count == 0)
            {
                snapshot.Errors.Add("Storage hardware not detected.");
                return;
            }

            foreach (var storage in storageHardware)
            {
                var temp = ReadTemperature(storage, new[] { "Temperature", "Drive Temperature", "HDD Temperature" });
                snapshot.Storage.Add(new StorageInfo
                {
                    Device = storage.Name,
                    Model = storage.Name,
                    TemperatureC = temp
                });

                if (!temp.HasValue)
                {
                    snapshot.Errors.Add($"Storage temperature not available for {storage.Name}.");
                }
            }
        }

        private static void PopulateFans(HardwareProbeSnapshot snapshot, Computer computer)
        {
            var fanSensors = computer.Hardware
                .SelectMany(GetAllSensors)
                .Where(sensor => sensor.SensorType == SensorType.Fan)
                .ToList();

            if (fanSensors.Count == 0)
            {
                snapshot.Errors.Add("Fan speed sensors not available.");
            }

            foreach (var sensor in fanSensors)
            {
                snapshot.Fans.Add(new FanInfo
                {
                    Name = sensor.Name,
                    Rpm = sensor.Value.HasValue ? (int)Math.Round(sensor.Value.Value) : null
                });
            }
        }

        private static double? ReadTemperature(IHardware? hardware, string[] preferredNames)
        {
            if (hardware == null) return null;
            return ReadSensorValue(hardware, SensorType.Temperature, preferredNames);
        }

        private static double? ReadSensorValue(IHardware hardware, SensorType sensorType, string[] preferredNames)
        {
            var sensors = GetAllSensors(hardware)
                .Where(sensor => sensor.SensorType == sensorType)
                .ToList();

            foreach (var name in preferredNames)
            {
                var match = sensors.FirstOrDefault(s => s.Name.Contains(name, StringComparison.OrdinalIgnoreCase));
                if (match?.Value != null)
                {
                    return Math.Round(match.Value.Value, 1);
                }
            }

            var fallback = sensors.FirstOrDefault(s => s.Value.HasValue);
            return fallback?.Value.HasValue == true ? Math.Round(fallback.Value.Value, 1) : null;
        }

        private static int? ReadMemorySensorMb(IHardware hardware, string[] preferredNames)
        {
            var sensors = GetAllSensors(hardware)
                .Where(sensor => sensor.SensorType == SensorType.SmallData || sensor.SensorType == SensorType.Data)
                .ToList();

            foreach (var name in preferredNames)
            {
                var match = sensors.FirstOrDefault(s => s.Name.Contains(name, StringComparison.OrdinalIgnoreCase));
                if (match?.Value != null)
                {
                    return (int)Math.Round(match.Value.Value);
                }
            }

            return null;
        }

        private static IEnumerable<ISensor> GetAllSensors(IHardware hardware)
        {
            foreach (var sensor in hardware.Sensors)
            {
                yield return sensor;
            }

            foreach (var subHardware in hardware.SubHardware)
            {
                foreach (var sensor in GetAllSensors(subHardware))
                {
                    yield return sensor;
                }
            }
        }

        private static (string Vendor, string Model, string Serial)? QueryBaseBoard()
        {
            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT Manufacturer, Product, SerialNumber FROM Win32_BaseBoard");
                foreach (var obj in searcher.Get().OfType<ManagementObject>())
                {
                    var vendor = obj["Manufacturer"]?.ToString();
                    var model = obj["Product"]?.ToString();
                    var serial = obj["SerialNumber"]?.ToString();
                    var vendorValue = vendor ?? string.Empty;
                    var modelValue = model ?? string.Empty;
                    var serialValue = serial ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(vendorValue) ||
                        !string.IsNullOrWhiteSpace(modelValue) ||
                        !string.IsNullOrWhiteSpace(serialValue))
                    {
                        return (vendorValue, modelValue, serialValue);
                    }
                }
            }
            catch
            {
                // Ignore WMI failure
            }

            return null;
        }

        private sealed class UpdateVisitor : IVisitor
        {
            public void VisitComputer(IComputer computer)
            {
                computer.Traverse(this);
            }

            public void VisitHardware(IHardware hardware)
            {
                hardware.Update();
                foreach (var subHardware in hardware.SubHardware)
                {
                    subHardware.Accept(this);
                }
            }

            public void VisitSensor(ISensor sensor)
            {
            }

            public void VisitParameter(IParameter parameter)
            {
            }
        }
    }
}
