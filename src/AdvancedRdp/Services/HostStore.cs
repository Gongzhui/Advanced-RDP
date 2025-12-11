using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using AdvancedRdp.Models;

namespace AdvancedRdp.Services;

public class HostStore
{
    private readonly string _configPath;
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public HostStore()
    {
        var appDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AdvancedRdp");
        Directory.CreateDirectory(appDir);
        _configPath = Path.Combine(appDir, "hosts.json");
    }

    public IEnumerable<HostEntry> LoadAll()
    {
        if (!File.Exists(_configPath)) yield break;

        var json = File.ReadAllText(_configPath);
        var hosts = JsonSerializer.Deserialize<List<HostEntry>>(json, JsonOptions);
        if (hosts == null) yield break;

        foreach (var host in hosts)
        {
            yield return host;
        }
    }

    public void SaveAll(IEnumerable<HostEntry> hosts)
    {
        var json = JsonSerializer.Serialize(hosts, JsonOptions);
        File.WriteAllText(_configPath, json);
    }
}
