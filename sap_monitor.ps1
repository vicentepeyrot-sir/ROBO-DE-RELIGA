Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class SapCloser {
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public const uint WM_CLOSE = 0x0010;
    public const uint WM_COMMAND = 0x0111;
    public const int IDOK = 1;

    public static int FecharAvisos() {
        int count = 0;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            if (IsWindowVisible(hWnd)) {
                StringBuilder sb = new StringBuilder(256);
                GetWindowText(hWnd, sb, 256);
                string title = sb.ToString().ToLower();
                
                // Se a janela principal se chama Informação, Aviso, etc
                if (title.Contains("informa") || title.Contains("aviso") || title.Contains("licen")) {
                    
                    bool isLicenca = false;

                    // Vasculha todos os textos DENTRO da janela para ver do que se trata
                    EnumChildWindows(hWnd, delegate(IntPtr hChild, IntPtr lChildParam) {
                        StringBuilder childSb = new StringBuilder(256);
                        GetWindowText(hChild, childSb, 256);
                        string childText = childSb.ToString().ToLower();

                        // Se encontrou a palavra licença ou expiração no texto interno
                        if (childText.Contains("expira") || childText.Contains("licen") || childText.Contains("dias")) {
                            isLicenca = true;
                            return false; // para de procurar nos textos internos
                        }
                        return true;
                    }, IntPtr.Zero);

                    // Só fecha se confirmou que é a janela de licença!
                    if (isLicenca) {
                        PostMessage(hWnd, WM_COMMAND, (IntPtr)IDOK, IntPtr.Zero);
                        PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                        count++;
                    }
                }
            }
            return true;
        }, IntPtr.Zero);
        return count;
    }
}
"@

Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " Monitor SAP Seguro (Lê texto interno) Iniciado!" -ForegroundColor Cyan
Write-Host " Pode deixar rodando no fundo. Ctrl+C para parar" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

$totalFechados = 0

while ($true) {
    $fechadosAgora = [SapCloser]::FecharAvisos()
    
    if ($fechadosAgora -gt 0) {
        $totalFechados += $fechadosAgora
        $hora = Get-Date -Format "HH:mm:ss"
        Write-Host "[$hora] Janela de licença validada e fechada com segurança! Total de hoje: $totalFechados" -ForegroundColor Green
    }
    
    # Aguarda 0.8s antes de checar de novo
    Start-Sleep -Milliseconds 800
}
