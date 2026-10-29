' rot_scan.ps1 — dsh-plugin-coreldraw 自带的只读诊断脚本
'
' 枚举 Running Object Table (ROT)，报告每个条目的 moniker 显示名，
' 并标记其中看起来属于 CorelDRAW 的条目。
'
' 兼容性：Windows 7 自带的 Windows PowerShell 2.0（CLR 2.0 / 无 C# 3 特性）。
' 输出协议（stdout，UTF-16LE）：
'   ROT|ENTRY|<显示名>|corel|other
'   ROT|CORELCOUNT|<n>
'   ROT|ERROR|<原因>
'
' 该脚本只读：它不创建任何 COM 对象、不写注册表、不启动任何进程。
'
' 输出编码：PowerShell 5.x 重定向到管道时默认 UTF-16LE；这里显式把
' [Console]::OutputEncoding 也设成 Unicode，使 Add-Type 编译出的 C#
' Console.WriteLine 与 PS 管道同为 UTF-16LE，Node 端按 utf16le 解码。
[Console]::OutputEncoding = [System.Text.Encoding]::Unicode

$code = @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class DshRotScan
{
    [DllImport("ole32.dll")]
    static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

    [DllImport("ole32.dll")]
    static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    static bool LooksCorel(string s)
    {
        return s != null && s.IndexOf("Corel", 0, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    public static int Scan()
    {
        IRunningObjectTable rot;
        int hr = GetRunningObjectTable(0, out rot);
        if (hr != 0 || rot == null)
        {
            Console.WriteLine("ROT|ERROR|GetRunningObjectTable failed hr=" + hr);
            return 1;
        }
        IEnumMoniker penum;
        rot.EnumRunning(out penum);
        if (penum == null)
        {
            Console.WriteLine("ROT|ERROR|EnumRunning null");
            return 1;
        }
        IBindCtx ctx;
        CreateBindCtx(0, out ctx);
        IMoniker[] arr = new IMoniker[1];
        int corelCount = 0;
        while (penum.Next(1, arr, IntPtr.Zero) == 0 && arr[0] != null)
        {
            string dn = null;
            try { arr[0].GetDisplayName(ctx, null, out dn); } catch { }
            if (dn == null) { dn = "<no-display-name>"; }
            bool corel = LooksCorel(dn);
            if (corel) corelCount++;
            Console.WriteLine("ROT|ENTRY|" + dn + "|" + (corel ? "corel" : "other"));
            arr = new IMoniker[1];
        }
        Console.WriteLine("ROT|CORELCOUNT|" + corelCount);
        return 0;
    }
}
'@

try
{
    Add-Type -TypeDefinition $code -Language CSharp
}
catch
{
    [string] $msg = $_.Exception.Message
    Write-Output ("ROT|ERROR|Add-Type failed: " + $msg)
    exit 1
}

try
{
    [DshRotScan]::Scan() | Out-Null
}
catch
{
    [string] $msg = $_.Exception.Message
    Write-Output ("ROT|ERROR|Scan threw: " + $msg)
    exit 1
}
