using System;
using System.Windows.Forms;

namespace AdvancedRdp.Controls;

public class RdpAxHost : AxHost
{
    public RdpAxHost(string clsid) : base(clsid)
    {
    }

    public object? OcxInstance => GetOcx();
}
