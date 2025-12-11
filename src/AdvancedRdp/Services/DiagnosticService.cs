using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;

namespace AdvancedRdp.Services;

public class DiagnosticService
{
    public DiagnosticResult Run(string host, int port)
    {
        var lines = new List<string>();
        var success = true;

        IPAddress[]? addresses = null;
        try
        {
            addresses = Dns.GetHostAddresses(host);
            lines.Add($"解析成功: {host} -> {string.Join(", ", addresses.Select(a => a.ToString()))}");
        }
        catch (Exception ex)
        {
            lines.Add($"解析失败: {ex.Message}");
            success = false;
        }

        if (addresses is { Length: > 0 })
        {
            var targetIp = addresses.First();
            try
            {
                using var ping = new Ping();
                var reply = ping.Send(targetIp, 2000);
                if (reply.Status == IPStatus.Success)
                {
                    lines.Add($"Ping 成功: {targetIp} 往返 {reply.RoundtripTime} ms");
                }
                else
                {
                    lines.Add($"Ping 失败: {targetIp} 状态 {reply.Status}");
                    success = false;
                }
            }
            catch (Exception ex)
            {
                lines.Add($"Ping 异常: {ex.Message}");
                success = false;
            }
        }

        if (addresses is { Length: > 0 })
        {
            var targetIp = addresses.First();
            using var client = new TcpClient();
            try
            {
                var connectTask = client.ConnectAsync(targetIp, port);
                var completed = connectTask.Wait(TimeSpan.FromSeconds(3));
                if (completed && client.Connected)
                {
                    lines.Add($"TCP 连接成功: {targetIp}:{port}");
                }
                else
                {
                    lines.Add($"TCP 连接超时/失败: {targetIp}:{port}");
                    success = false;
                }
            }
            catch (SocketException ex)
            {
                lines.Add($"TCP 连接被拒绝/失败: {targetIp}:{port} ({ex.SocketErrorCode})");
                success = false;
            }
            catch (Exception ex)
            {
                lines.Add($"TCP 测试异常: {ex.Message}");
                success = false;
            }
        }

        if (!success)
        {
            lines.Add($"建议检查：1) 主机名/IP 是否正确；2) 目标 {port} 端口是否开放/防火墙策略；3) 凭据是否正确；4) VPN/内网可达性。");
        }

        return new DiagnosticResult(success, lines);
    }
}

public record DiagnosticResult(bool Success, IReadOnlyList<string> Lines);
