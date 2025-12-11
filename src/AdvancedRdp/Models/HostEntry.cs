using System.Text.Json.Serialization;

namespace AdvancedRdp.Models;

public class HostEntry
{
    public string Name { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string Username { get; set; } = string.Empty;
    public string Domain { get; set; } = string.Empty;
    public int DesktopWidth { get; set; } = 1280;
    public int DesktopHeight { get; set; } = 720;
    public bool RedirectClipboard { get; set; } = true;
    public bool FullScreen { get; set; } = true;

    [JsonIgnore]
    public string CredentialKey => $"advancedrdp:{Name}";
}
