<#
.SYNOPSIS
    Aplicativo de Automacao Microsoft Teams com Interface Grafica (WPF/PowerShell)
    Permite selecionar planilha Excel, mapear colunas dinamicamente e enviar
    mensagens automaticas no Microsoft Teams desktop sem necessidade de Azure AD.
    Inclui tela de resumo com detalhes de enviados, falhas e motivos, e exportacao CSV.
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing, Microsoft.VisualBasic

[System.Windows.Forms.Application]::EnableVisualStyles()

# Helper nativo Win32 para foco confiavel, maximizacao e clique no Teams
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class TeamsHelper {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

    public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    public const uint MOUSEEVENTF_LEFTUP = 0x04;
    public const int SW_MAXIMIZE = 3;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static IntPtr ObterJanelaTeams() {
        IntPtr foundHwnd = IntPtr.Zero;
        EnumWindows((hWnd, lParam) => {
            if (IsWindowVisible(hWnd)) {
                RECT rect;
                if (GetWindowRect(hWnd, out rect)) {
                    int w = rect.Right - rect.Left;
                    int h = rect.Bottom - rect.Top;
                    if (w > 500 && h > 300) {
                        uint pid = 0;
                        GetWindowThreadProcessId(hWnd, out pid);
                        try {
                            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById((int)pid);
                            if (p.ProcessName.IndexOf("teams", StringComparison.OrdinalIgnoreCase) >= 0) {
                                foundHwnd = hWnd;
                                return false; // Encontrou a janela principal do Teams
                            }
                        } catch {}
                    }
                }
            }
            return true;
        }, IntPtr.Zero);
        return foundHwnd;
    }

    public static void AtivarEMaximizar(IntPtr hwnd) {
        if (hwnd == IntPtr.Zero) {
            hwnd = ObterJanelaTeams();
        }
        if (hwnd != IntPtr.Zero) {
            ShowWindow(hwnd, SW_MAXIMIZE);
            keybd_event(0x12, 0, 0, 0); // ALT Down
            keybd_event(0x12, 0, 2, 0); // ALT Up
            SetForegroundWindow(hwnd);
            ShowWindow(hwnd, SW_MAXIMIZE);
        }
    }

    public static void FocarJanela(IntPtr hwnd) {
        AtivarEMaximizar(hwnd);
    }

    public static void ClicarNaBarraPesquisa() {
        IntPtr hwnd = ObterJanelaTeams();
        if (hwnd == IntPtr.Zero) {
            hwnd = GetForegroundWindow();
        }
        RECT rect;
        if (GetWindowRect(hwnd, out rect)) {
            int w = rect.Right - rect.Left;
            // A barra "Pesquisar (Ctrl+E)" fica centralizada no topo e a ~24px abaixo da borda superior
            int targetX = rect.Left + (w / 2);
            int targetY = rect.Top + 24;
            SetCursorPos(targetX, targetY);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }
    }

    public static void ClicarNoCampoMensagem() {
        IntPtr hwnd = ObterJanelaTeams();
        if (hwnd == IntPtr.Zero) {
            hwnd = GetForegroundWindow();
        }
        RECT rect;
        if (GetWindowRect(hwnd, out rect)) {
            int w = rect.Right - rect.Left;
            int h = rect.Bottom - rect.Top;
            // O campo "Digite uma mensagem" fica no painel de chat e a ~75px do fundo da janela (bem acima da barra de icones)
            int targetX = (w > 800) ? (rect.Left + 360 + (int)((w - 360) * 0.45)) : (rect.Left + (int)(w * 0.60));
            int targetY = rect.Bottom - 75;
            SetCursorPos(targetX, targetY);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }
    }

    public static void ClicarNoBotaoEnviar() {
        IntPtr hwnd = ObterJanelaTeams();
        if (hwnd == IntPtr.Zero) {
            hwnd = GetForegroundWindow();
        }
        RECT rect;
        if (GetWindowRect(hwnd, out rect)) {
            // O botao Enviar (icone de aviao/seta) fica no canto inferior direito da caixa de chat
            int targetX = rect.Right - 38;
            int targetY = rect.Bottom - 45;
            SetCursorPos(targetX, targetY);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        }
    }
}
"@

# Estrutura XAML da Interface Principal
[xml]$xamlPrincipal = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Automacao Microsoft Teams - Envio Automatico via Excel"
        Height="760" Width="1220" MinHeight="580" MinWidth="980"
        WindowState="Maximized" WindowStartupLocation="CenterScreen"
        Background="#F1F5F9" FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,6"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#1E293B"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="Foreground" Value="#334155"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="0,1,0,1"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
    </Window.Resources>

    <Grid Margin="10,6,10,10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- 0: Header -->
            <RowDefinition Height="Auto"/> <!-- 1: Selecao de Planilha -->
            <RowDefinition Height="Auto"/> <!-- 2: Mapeamento de Colunas -->
            <RowDefinition Height="Auto"/> <!-- 3: Configurador de Regras -->
            <RowDefinition Height="*"/>    <!-- 4: Previa e Logs -->
            <RowDefinition Height="Auto"/> <!-- 5: Botoes de Acao -->
        </Grid.RowDefinitions>

        <!-- CABECALHO -->
        <Border Grid.Row="0" Background="#1E3A8A" CornerRadius="6" Padding="12,7" Margin="0,0,0,6">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Automacao Microsoft Teams - Envio Automatico via Excel" FontSize="16" FontWeight="Bold" Foreground="White"/>
                    <TextBlock Text="Disparo no Teams nativo com leitura dinamica, configurador de 2a mensagem, protecao contra erros e relatorio final" FontSize="11" Foreground="#93C5FD" Margin="0,2,0,0"/>
                </StackPanel>
                <TextBlock Grid.Column="1" Text="Versao PowerShell RPA" FontSize="11" Foreground="#BFDBFE" VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- 1. SELECAO DE PLANILHA -->
        <GroupBox Grid.Row="1" Header="1. Selecao da Planilha Excel">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="180"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="TxtCaminhoPlanilha" Grid.Column="0" IsReadOnly="True" Margin="0,0,8,0" Text="Selecione um arquivo .xlsx ou .xlsm..."/>
                <Button x:Name="BtnProcurar" Grid.Column="1" Content="Procurar..." Padding="12,3" Margin="0,0,10,0" Background="#2563EB" Foreground="White" FontWeight="SemiBold"/>
                
                <StackPanel Grid.Column="2" Orientation="Vertical">
                    <Label Content="Aba do Excel:"/>
                    <ComboBox x:Name="CmbAbas"/>
                </StackPanel>
                <Button x:Name="BtnCarregarColunas" Grid.Column="3" Content="Carregar e Mapear" Padding="12,3" Margin="8,14,0,0" Background="#059669" Foreground="White" FontWeight="SemiBold"/>
            </Grid>
        </GroupBox>

        <!-- 2. MAPEAMENTO DE COLUNAS -->
        <GroupBox Grid.Row="2" Header="2. Mapeamento das Colunas da Planilha (Ajuste se os nomes das colunas mudarem)">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Coluna 1 -->
                <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,8,4">
                    <Label Content="Matricula / ID:"/>
                    <ComboBox x:Name="CmbColMatricula"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,0,8,0">
                    <Label Content="Nome do Usuario:"/>
                    <ComboBox x:Name="CmbColNome"/>
                </StackPanel>

                <!-- Coluna 2 -->
                <StackPanel Grid.Row="0" Grid.Column="1" Margin="0,0,8,4">
                    <Label Content="Mensagem (1o Contato):"/>
                    <ComboBox x:Name="CmbColMensagem1"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Grid.Column="1" Margin="0,0,8,0">
                    <Label Content="Mensagem (2o Contato):"/>
                    <ComboBox x:Name="CmbColMensagem2"/>
                </StackPanel>

                <!-- Coluna 3 -->
                <StackPanel Grid.Row="0" Grid.Column="2" Margin="0,0,8,4">
                    <Label Content="Status do Envio:"/>
                    <ComboBox x:Name="CmbColStatus"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Grid.Column="2" Margin="0,0,8,0">
                    <Label Content="Numero de Contatos Atual:"/>
                    <ComboBox x:Name="CmbColNumero"/>
                </StackPanel>

                <!-- Coluna 4 -->
                <StackPanel Grid.Row="0" Grid.Column="3" Margin="0,0,8,4">
                    <Label Content="Ultima Data de Contato:"/>
                    <ComboBox x:Name="CmbColUltimaData"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Grid.Column="3" Margin="0,0,8,0">
                    <Label Content="E-mail Teams Direto (opcional):"/>
                    <ComboBox x:Name="CmbColEmail"/>
                </StackPanel>

                <!-- Coluna 5: Analista / Quem enviou -->
                <StackPanel Grid.Row="0" Grid.Column="4" Margin="0,0,0,4">
                    <Label Content="Coluna Analista (Planilha):" FontWeight="SemiBold"/>
                    <ComboBox x:Name="CmbColAnalista" ToolTip="Selecione a coluna da planilha Excel onde será gravado quem enviou a mensagem (ex: ANALISTA)"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Grid.Column="4" Margin="0,0,0,0">
                    <Label Content="Nome do Analista (Registro):" FontWeight="SemiBold"/>
                    <TextBox x:Name="TxtNomeAnalista" ToolTip="Nome ou matricula do analista gravado na planilha a cada envio com sucesso"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- 3. CONFIGURADOR DE LIMITES, 2A MENSAGEM E REGRAS -->
        <GroupBox Grid.Row="3" Header="3. Configurador de Limites, 2a Mensagem e Formato de Contato">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="130"/>
                    <ColumnDefinition Width="145"/>
                    <ColumnDefinition Width="220"/>
                    <ColumnDefinition Width="105"/>
                    <ColumnDefinition Width="115"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Margin="0,0,8,0">
                    <Label Content="Max. contatos p/ pessoa:" FontWeight="SemiBold"/>
                    <TextBox x:Name="TxtMaxContatos" Text="2" ToolTip="Ex: 1 = Apenas 1 mensagem. 2 = Permite enviar 1a e 2a mensagem."/>
                </StackPanel>

                <StackPanel Grid.Column="1" Margin="0,0,8,0">
                    <Label Content="Dias de espera p/ 2a msg:" FontWeight="SemiBold"/>
                    <TextBox x:Name="TxtDiasSegundo" Text="3" ToolTip="Quantos dias apos o 1o contato a pessoa estara liberada para receber a 2a mensagem."/>
                </StackPanel>

                <StackPanel Grid.Column="2" Margin="0,0,8,0">
                    <Label Content="Formato do Teams (ex: {matricula}):"/>
                    <TextBox x:Name="TxtFormatoEmail" ToolTip="Digite o dominio real do Teams da sua empresa ou use apenas {matricula} se o Teams busca direto pela matricula"/>
                </StackPanel>

                <StackPanel Grid.Column="3" Margin="0,0,8,0">
                    <Label Content="Espera Teams (s):"/>
                    <TextBox x:Name="TxtEsperaTeams" Text="4"/>
                </StackPanel>

                <StackPanel Grid.Column="4" Margin="0,0,8,0">
                    <Label Content="Intervalo envios (s):"/>
                    <TextBox x:Name="TxtIntervaloEnvio" Text="3"/>
                </StackPanel>

                <Button x:Name="BtnRecalcular" Grid.Column="5" Content="Aplicar Regras" Padding="12,3" Margin="4,14,4,0" Background="#4F46E5" Foreground="White" FontWeight="SemiBold" ToolTip="Recalcula a previa na hora com os novos limites"/>

                <Button x:Name="BtnSalvarConfig" Grid.Column="6" Padding="10,3" Margin="4,14,8,0" Background="#0284C7" Foreground="White" FontWeight="SemiBold" ToolTip="Salva a planilha, abas, colunas e configuracoes para abrir tudo automaticamente na proxima vez">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <Viewbox Width="14" Height="14" Margin="0,0,6,0">
                            <Canvas Width="24" Height="24">
                                <Path Fill="White" Data="M19,21H5C3.89,21 3,20.1 3,19V5C3,3.89 3.89,3 5,3H16L21,8V19C21,20.1 20.1,21 19,21M17,19V11H7V19H17M15,5H5V9H15V5Z"/>
                            </Canvas>
                        </Viewbox>
                        <TextBlock Text="Salvar Config" VerticalAlignment="Center"/>
                    </StackPanel>
                </Button>

                <StackPanel Grid.Column="7" VerticalAlignment="Center" Margin="4,12,0,0">
                    <CheckBox x:Name="ChkConfirmarEnvio" Content="Confirmar antes de gravar no Excel" FontWeight="SemiBold" Foreground="#1E40AF" ToolTip="Pergunta se a mensagem realmente foi enviada no Teams antes de atualizar a planilha"/>
                    <CheckBox x:Name="ChkModoSimulacao" Content="Modo Simulacao (nao envia)" FontWeight="Normal" Foreground="#B45309" Margin="0,2,0,0"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- 4. PREVIA DOS DADOS (COM ROLAGEM HORIZONTAL/VERTICAL E LOGS) -->
        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="3.2*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="1.2*"/>
            </Grid.ColumnDefinitions>

            <!-- TABELA DE PREVIA COMPLETA -->
            <GroupBox Grid.Column="0" Margin="0,0,4,0">
                <GroupBox.Header>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Previa dos Contatos" FontWeight="Bold"/>
                        <TextBlock x:Name="TxtContadores" Text=" (Nenhuma planilha carregada)" Foreground="#64748B" FontWeight="Normal" Margin="6,0,0,0"/>
                    </StackPanel>
                </GroupBox.Header>
                <DataGrid x:Name="GridPrevia" AutoGenerateColumns="False" IsReadOnly="True"
                          HeadersVisibility="Column" GridLinesVisibility="All"
                          HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                          ScrollViewer.HorizontalScrollBarVisibility="Auto"
                          ScrollViewer.VerticalScrollBarVisibility="Auto"
                          ScrollViewer.CanContentScroll="False"
                          Background="White" RowHeaderWidth="0" FontSize="11">
                    <DataGrid.Resources>
                        <Style TargetType="DataGridCell">
                            <Setter Property="Padding" Value="4,4"/>
                            <Setter Property="VerticalContentAlignment" Value="Top"/>
                        </Style>
                    </DataGrid.Resources>
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Linha" Binding="{Binding Linha}" Width="45"/>
                        <DataGridTextColumn Header="Matricula" Binding="{Binding Matricula}" Width="85"/>
                        <DataGridTextColumn Header="Nome" Binding="{Binding Nome}" Width="135"/>
                        <DataGridTextColumn Header="Destinatario" Binding="{Binding Destinatario}" Width="150"/>
                        <DataGridTextColumn Header="Contatos" Binding="{Binding Numero}" Width="65"/>
                        <DataGridTextColumn Header="Msg a Enviar" Binding="{Binding TipoMensagem}" Width="110"/>
                        <DataGridTextColumn Header="Situacao" Binding="{Binding Situacao}" Width="145"/>
                        <DataGridTextColumn Header="Status Atual" Binding="{Binding Status}" Width="110"/>
                        <DataGridTextColumn Header="Ultima Data" Binding="{Binding UltimaData}" Width="130"/>
                        
                        <!-- 1a Mensagem completa com quebra de linha (nao corta o texto) -->
                        <DataGridTextColumn Header="1a Mensagem (Planilha)" Binding="{Binding Mensagem1}" Width="340">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="TextWrapping" Value="Wrap"/>
                                    <Setter Property="VerticalAlignment" Value="Top"/>
                                    <Setter Property="Padding" Value="2,2"/>
                                    <Setter Property="Foreground" Value="#1E293B"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>

                        <!-- 2a Mensagem completa com quebra de linha (nao corta o texto) -->
                        <DataGridTextColumn Header="2a Mensagem (Follow-up)" Binding="{Binding Mensagem2}" Width="340">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="TextWrapping" Value="Wrap"/>
                                    <Setter Property="VerticalAlignment" Value="Top"/>
                                    <Setter Property="Padding" Value="2,2"/>
                                    <Setter Property="Foreground" Value="#334155"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>

                        <!-- Mensagem que sera efetivamente disparada agora -->
                        <DataGridTextColumn Header="Texto que Sera Enviado" Binding="{Binding Mensagem}" Width="380">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="TextWrapping" Value="Wrap"/>
                                    <Setter Property="VerticalAlignment" Value="Top"/>
                                    <Setter Property="Padding" Value="2,2"/>
                                    <Setter Property="FontWeight" Value="SemiBold"/>
                                    <Setter Property="Foreground" Value="#0F172A"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>
                    </DataGrid.Columns>
                </DataGrid>
            </GroupBox>

            <!-- DIVISOR REDIMENSIONAVEL (GRID SPLITTER) -->
            <GridSplitter Grid.Column="1" Width="6" HorizontalAlignment="Center" VerticalAlignment="Stretch"
                          Background="#CBD5E1" Cursor="SizeWE" ToolTip="Arraste para redimensionar a tabela e os logs"/>

            <!-- LOGS EM TEMPO REAL -->
            <GroupBox Grid.Column="2" Margin="4,0,0,0" Header="Logs de Execucao em Tempo Real">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBox x:Name="TxtLogs" Grid.Row="0" IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                             Background="#0F172A" Foreground="#38BDF8" FontFamily="Consolas" FontSize="11"
                             Padding="6" TextWrapping="Wrap"/>
                    <ProgressBar x:Name="BarraProgresso" Grid.Row="1" Height="14" Margin="0,4,0,0" Minimum="0" Maximum="100" Value="0"/>
                </Grid>
            </GroupBox>
        </Grid>

        <!-- 5. PAINEL DE ACAO E BOTOES -->
        <Border Grid.Row="5" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Padding="8" Margin="0,6,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <Button x:Name="BtnIniciar" Grid.Column="0" Content="Iniciar Envio Automatico" Padding="16,6" Background="#16A34A" Foreground="White" FontWeight="Bold" FontSize="12" Margin="0,0,8,0"/>
                <Button x:Name="BtnTestarUm" Grid.Column="1" Content="Testar Apenas o 1o Contato" Padding="12,6" Background="#0284C7" Foreground="White" FontWeight="SemiBold" FontSize="11" Margin="0,0,8,0"/>
                <TextBlock x:Name="TxtStatusGeral" Grid.Column="2" Text="Pronto para carregar planilha." VerticalAlignment="Center" Margin="10,0" Foreground="#475467" FontSize="11"/>
                <Button x:Name="BtnVerUltimoResumo" Grid.Column="3" Content="Ver Ultimo Relatorio" Padding="12,6" Background="#0891B2" Foreground="White" FontWeight="SemiBold" FontSize="11" Margin="0,0,8,0" IsEnabled="False"/>
                <Button x:Name="BtnParar" Grid.Column="4" Content="Parar Envio" Padding="14,6" Background="#DC2626" Foreground="White" FontWeight="Bold" FontSize="11" IsEnabled="False"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Estrutura XAML da Janela de Resumo Final
[xml]$xamlResumo = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Relatorio Final de Envio - Resumo da Execucao"
        Height="600" Width="920" MinHeight="480" MinWidth="750"
        WindowStartupLocation="CenterOwner" Background="#F8FAFC"
        FontFamily="Segoe UI">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Cards Resumo -->
            <RowDefinition Height="*"/>    <!-- DataGrid de Resultados -->
            <RowDefinition Height="Auto"/> <!-- Botoes -->
        </Grid.RowDefinitions>

        <!-- Indicadores Resumo -->
        <Grid Grid.Row="0" Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Card Total -->
            <Border Grid.Column="0" Background="#EFF6FF" BorderBrush="#BFDBFE" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,8,0">
                <StackPanel>
                    <TextBlock Text="Total Processados" FontSize="11" Foreground="#1E40AF" FontWeight="SemiBold"/>
                    <TextBlock x:Name="TxtResumoTotal" Text="0" FontSize="26" FontWeight="Bold" Foreground="#1E3A8A" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>

            <!-- Card Sucesso -->
            <Border Grid.Column="1" Background="#ECFDF5" BorderBrush="#A7F3D0" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,8,0">
                <StackPanel>
                    <TextBlock Text="Enviados com Sucesso" FontSize="11" Foreground="#065F46" FontWeight="SemiBold"/>
                    <TextBlock x:Name="TxtResumoSucessos" Text="0" FontSize="26" FontWeight="Bold" Foreground="#047857" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>

            <!-- Card Erros / Falhas -->
            <Border Grid.Column="2" Background="#FEF2F2" BorderBrush="#FECACA" BorderThickness="1" CornerRadius="8" Padding="12">
                <StackPanel>
                    <TextBlock Text="Erros / Nao Enviados" FontSize="11" Foreground="#991B1B" FontWeight="SemiBold"/>
                    <TextBlock x:Name="TxtResumoErros" Text="0" FontSize="26" FontWeight="Bold" Foreground="#DC2626" Margin="0,2,0,0"/>
                </StackPanel>
            </Border>
        </Grid>

        <!-- Tabela Detalhada com Erros e Motivos -->
        <GroupBox Grid.Row="1" Header="Detalhamento dos Contatos, Resultados e Motivos de Falha" FontWeight="SemiBold">
            <DataGrid x:Name="GridResumo" AutoGenerateColumns="False" IsReadOnly="True"
                      HeadersVisibility="Column" GridLinesVisibility="All"
                      HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto"
                      ScrollViewer.HorizontalScrollBarVisibility="Auto"
                      ScrollViewer.VerticalScrollBarVisibility="Auto"
                      ScrollViewer.CanContentScroll="True"
                      Background="White" RowHeaderWidth="0" FontSize="11" Margin="0,6,0,0">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Linha" Binding="{Binding Linha}" Width="45"/>
                    <DataGridTextColumn Header="Matricula" Binding="{Binding Matricula}" Width="75"/>
                    <DataGridTextColumn Header="Nome" Binding="{Binding Nome}" Width="120"/>
                    <DataGridTextColumn Header="Destinatario" Binding="{Binding Destinatario}" Width="130"/>
                    <DataGridTextColumn Header="Analista" Binding="{Binding Analista}" Width="90"/>
                    <DataGridTextColumn Header="Tipo Msg" Binding="{Binding Tipo}" Width="85"/>
                    <DataGridTextColumn Header="Resultado" Binding="{Binding Resultado}" Width="95"/>
                    <DataGridTextColumn Header="Motivo / Detalhe do Envio" Binding="{Binding Motivo}" Width="*"/>
                </DataGrid.Columns>
            </DataGrid>
        </GroupBox>

        <!-- Barra Inferior de Botoes -->
        <Grid Grid.Row="2" Margin="0,12,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <TextBlock x:Name="TxtInfoResumo" Grid.Column="0" Text="Relatorio gerado com sucesso." VerticalAlignment="Center" Foreground="#64748B" FontSize="12"/>
            <Button x:Name="BtnExportarCsv" Grid.Column="1" Content="Exportar Relatorio (.csv)" Padding="14,8" Background="#0284C7" Foreground="White" FontWeight="SemiBold" Margin="0,0,8,0"/>
            <Button x:Name="BtnFecharResumo" Grid.Column="2" Content="Fechar" Padding="18,8" Background="#475569" Foreground="White" FontWeight="SemiBold"/>
        </Grid>
    </Grid>
</Window>
"@

# Carregar o XAML Principal no WPF
$readerPrincipal = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlPrincipal.OuterXml))
$window = [System.Windows.Markup.XamlReader]::Load($readerPrincipal)

# Garantir que a janela nunca abra acima do topo do monitor
$window.Add_Loaded({
        if ($window.Top -lt 0) { $window.Top = 0 }
        if ($window.Left -lt 0) { $window.Left = 0 }
    })

# Mapear controles da Janela Principal
$TxtCaminhoPlanilha = $window.FindName("TxtCaminhoPlanilha")
$BtnProcurar = $window.FindName("BtnProcurar")
$CmbAbas = $window.FindName("CmbAbas")
$BtnCarregarColunas = $window.FindName("BtnCarregarColunas")

$CmbColMatricula = $window.FindName("CmbColMatricula")
$CmbColNome = $window.FindName("CmbColNome")
$CmbColMensagem1 = $window.FindName("CmbColMensagem1")
$CmbColMensagem2 = $window.FindName("CmbColMensagem2")
$CmbColStatus = $window.FindName("CmbColStatus")
$CmbColNumero = $window.FindName("CmbColNumero")
$CmbColUltimaData = $window.FindName("CmbColUltimaData")
$CmbColEmail = $window.FindName("CmbColEmail")
$CmbColAnalista = $window.FindName("CmbColAnalista")
$TxtNomeAnalista = $window.FindName("TxtNomeAnalista")

$TxtMaxContatos = $window.FindName("TxtMaxContatos")
$TxtDiasSegundo = $window.FindName("TxtDiasSegundo")
$BtnRecalcular = $window.FindName("BtnRecalcular")
$BtnSalvarConfig = $window.FindName("BtnSalvarConfig")
$TxtFormatoEmail = $window.FindName("TxtFormatoEmail")
$TxtFormatoEmail.Text = "{matricula}" # Padrao: busca direto pela matricula ou usuario personaliza
$TxtEsperaTeams = $window.FindName("TxtEsperaTeams")
$TxtIntervaloEnvio = $window.FindName("TxtIntervaloEnvio")
$ChkConfirmarEnvio = $window.FindName("ChkConfirmarEnvio")
$ChkModoSimulacao = $window.FindName("ChkModoSimulacao")

$GridPrevia = $window.FindName("GridPrevia")
$TxtContadores = $window.FindName("TxtContadores")
$TxtLogs = $window.FindName("TxtLogs")
$BarraProgresso = $window.FindName("BarraProgresso")

$BtnIniciar = $window.FindName("BtnIniciar")
$BtnTestarUm = $window.FindName("BtnTestarUm")
$BtnVerUltimoResumo = $window.FindName("BtnVerUltimoResumo")
$BtnParar = $window.FindName("BtnParar")
$TxtStatusGeral = $window.FindName("TxtStatusGeral")

# Variaveis globais de controle e historico de relatorio
$global:PararSolicitado = $false
$global:ColunasDetectadas = @()
$global:ListaContatos = @()
$global:RelatorioFinal = [System.Collections.ArrayList]::new()
$global:PastaApp = if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot "App.ps1"))) { $PSScriptRoot } else { "W:\N8N\SISTEMA BANCO\prototipo_teams_powershell" }

# Funcao auxiliar para registrar logs com data/hora e scroll automatico
function Log-Msg($msg, $destaque = $false) {
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $linha = "[$timestamp] $msg`r`n"
    if ($destaque) {
        $linha = "----------------------------------------`r`n$linha----------------------------------------`r`n"
    }
    $TxtLogs.AppendText($linha)
    $TxtLogs.ScrollToEnd()
    [System.Windows.Forms.Application]::DoEvents()
}

# Normalizar texto para comparacoes
function Normalizar-Texto($txt) {
    if ([string]::IsNullOrWhiteSpace($txt)) { return "" }
    return ($txt.Trim().ToLower() -replace '[^a-z0-9]', '')
}

# Obter nome ou display name do analista/usuario logado no Windows/Teams/AD
function Obter-Nome-Analista-Padrao {
    $fullName = ""
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction SilentlyContinue
        $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current
        if ($user -and $user.DisplayName) { $fullName = $user.DisplayName.Trim() }
    } catch {}

    if ([string]::IsNullOrWhiteSpace($fullName)) {
        try {
            $adsi = [adsi]"WinNT://$env:USERDOMAIN/$env:USERNAME,user"
            if ($adsi.FullName) { $fullName = [string]$adsi.FullName.ToString().Trim() }
        } catch {}
    }

    if ([string]::IsNullOrWhiteSpace($fullName)) {
        try {
            $wmi = Get-CimInstance Win32_UserAccount -Filter "Name='$($env:USERNAME)' and Domain='$($env:USERDOMAIN)'" -ErrorAction SilentlyContinue
            if ($wmi -and $wmi.FullName) { $fullName = $wmi.FullName.Trim() }
        } catch {}
    }

    if ([string]::IsNullOrWhiteSpace($fullName)) {
        $fullName = if ($env:USERNAME) { $env:USERNAME.Trim() } else { "Analista" }
    }
    return $fullName
}

# Inicializar o campo do analista com o usuario detectado
if ($TxtNomeAnalista) {
    $TxtNomeAnalista.Text = Obter-Nome-Analista-Padrao
}

# Procurar Planilha via Dialogo nativo
$BtnProcurar.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = "Planilhas Excel (*.xlsx; *.xlsm)|*.xlsx;*.xlsm|Todos os Arquivos (*.*)|*.*"
        $dialog.Title = "Selecione a Planilha de Contatos"
    
        $pastaPadrao = "W:\N8N\SISTEMA BANCO\prototipo_teams_manual\dados"
        if (Test-Path $pastaPadrao) { $dialog.InitialDirectory = $pastaPadrao }

        if ($dialog.ShowDialog() -eq $true) {
            $TxtCaminhoPlanilha.Text = $dialog.FileName
            Carregar-Abas-Planilha $dialog.FileName
        }
    })

# Ler Abas do arquivo Excel selecionado
function Carregar-Abas-Planilha($caminho, $abaPreSelecionada = $null) {
    Log-Msg "Inspecionando abas do arquivo: $(Split-Path $caminho -Leaf)..."
    $CmbAbas.Items.Clear()

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open($caminho, $null, $true)
        foreach ($sh in $wb.Sheets) {
            $CmbAbas.Items.Add($sh.Name) | Out-Null
        }
        $wb.Close($false)
        if ($CmbAbas.Items.Count -gt 0) {
            if ($abaPreSelecionada -and $CmbAbas.Items.Contains($abaPreSelecionada)) {
                $CmbAbas.SelectedItem = $abaPreSelecionada
            }
            else {
                $CmbAbas.SelectedIndex = 0
            }
            Log-Msg "Abas detectadas: $($CmbAbas.Items.Count)."
        }
    }
    catch {
        Log-Msg "Erro ao ler abas do Excel: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Nao foi possivel ler o arquivo: $($_.Exception.Message)", "Erro no Excel", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
    finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
    }
}

# Preencher ComboBox com as colunas detectadas e auto-mapear (priorizando configuracao salva)
function Popular-Combo-Colunas($combo, $colunas, $nomesPadrao, $obrigatorio = $true, $salvo = $null) {
    $combo.Items.Clear()
    if (-not $obrigatorio) {
        $combo.Items.Add("(Nenhum / Desativado)") | Out-Null
    }

    $indiceSelecionado = -1
    $index = if ($obrigatorio) { 0 } else { 1 }

    foreach ($col in $colunas) {
        $combo.Items.Add($col.Nome) | Out-Null

        # 1. Se houver valor salvo previamente, tem prioridade absoluta
        if (-not [string]::IsNullOrWhiteSpace($salvo) -and $col.Nome.Trim().ToLower() -eq $salvo.Trim().ToLower()) {
            $indiceSelecionado = $index
        }

        # 2. Caso nao tenha salvo, usa busca por padrao de texto
        if ($indiceSelecionado -eq -1) {
            $colNorm = Normalizar-Texto $col.Nome
            foreach ($padrao in $nomesPadrao) {
                $padraoNorm = Normalizar-Texto $padrao
                if ($colNorm -like "*$padraoNorm*") {
                    $indiceSelecionado = $index
                    break
                }
            }
        }
        $index++
    }

    if ($indiceSelecionado -ge 0) {
        $combo.SelectedIndex = $indiceSelecionado
    }
    elseif ($obrigatorio -and $combo.Items.Count -gt 0) {
        $combo.SelectedIndex = 0
    }
    elseif (-not $obrigatorio) {
        $combo.SelectedIndex = 0
    }
}

# Funcao para encontrar o numero da coluna a partir do nome selecionado
function Obter-Indice-Coluna($nomeColuna) {
    if ([string]::IsNullOrWhiteSpace($nomeColuna) -or $nomeColuna -like "*(Nenhum*") { return 0 }
    $match = $global:ColunasDetectadas | Where-Object { $_.Nome -eq $nomeColuna } | Select-Object -First 1
    if ($match) { return $match.Numero }
    return 0
}

# Salvar Configuracoes no arquivo config.json
function Salvar-Configuracoes {
    $cfgPath = Join-Path $global:PastaApp "config.json"
    $config = [PSCustomObject]@{
        CaminhoPlanilha = $TxtCaminhoPlanilha.Text
        AbaSelecionada  = $CmbAbas.Text
        ColMatricula    = $CmbColMatricula.Text
        ColNome         = $CmbColNome.Text
        ColMensagem1    = $CmbColMensagem1.Text
        ColMensagem2    = $CmbColMensagem2.Text
        ColStatus       = $CmbColStatus.Text
        ColNumero       = $CmbColNumero.Text
        ColUltimaData   = $CmbColUltimaData.Text
        ColEmail        = $CmbColEmail.Text
        ColAnalista     = $CmbColAnalista.Text
        NomeAnalista    = $TxtNomeAnalista.Text
        MaxContatos     = $TxtMaxContatos.Text
        DiasSegundo     = $TxtDiasSegundo.Text
        FormatoEmail    = $TxtFormatoEmail.Text
        EsperaTeams     = $TxtEsperaTeams.Text
        IntervaloEnvio  = $TxtIntervaloEnvio.Text
        ConfirmarEnvio  = [bool]$ChkConfirmarEnvio.IsChecked
        ModoSimulacao   = [bool]$ChkModoSimulacao.IsChecked
    }
    try {
        $json = $config | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($cfgPath, $json, [System.Text.Encoding]::UTF8)
        Log-Msg "Configuracoes salvas automaticamente em config.json."
    }
    catch {
        Log-Msg "Aviso ao salvar config.json: $($_.Exception.Message)"
    }
}

# Carregar Dados e Atualizar Tabela de Previa
function Carregar-Previa-Dados {
    $caminho = $TxtCaminhoPlanilha.Text
    $aba = $CmbAbas.Text

    if (-not (Test-Path $caminho)) { return }

    $colMat = Obter-Indice-Coluna $CmbColMatricula.Text
    $colNome = Obter-Indice-Coluna $CmbColNome.Text
    $colMsg1 = Obter-Indice-Coluna $CmbColMensagem1.Text
    $colMsg2 = Obter-Indice-Coluna $CmbColMensagem2.Text
    $colStatus = Obter-Indice-Coluna $CmbColStatus.Text
    $colNum = Obter-Indice-Coluna $CmbColNumero.Text
    $colData = Obter-Indice-Coluna $CmbColUltimaData.Text
    $colEmail = Obter-Indice-Coluna $CmbColEmail.Text

    $formatoEmail = $TxtFormatoEmail.Text
    if ([string]::IsNullOrWhiteSpace($formatoEmail)) { $formatoEmail = "{matricula}" }

    $maxContatos = 2
    [int]::TryParse($TxtMaxContatos.Text, [ref]$maxContatos) | Out-Null
    $diasSegundo = 3
    [int]::TryParse($TxtDiasSegundo.Text, [ref]$diasSegundo) | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open($caminho, $null, $true)
        $excel.Calculation = -4105 # xlCalculationAutomatic
        try { $wb.ForceFullCalculation = $true } catch {}
        $excel.Calculate()
        $excel.CalculateFull()
        $ws = if ($aba) { $wb.Sheets.Item($aba) } else { $wb.Sheets.Item(1) }
        $totalLinhas = $ws.UsedRange.Rows.Count

        $global:ListaContatos = [System.Collections.ArrayList]::new()

        for ($linha = 2; $linha -le $totalLinhas; $linha++) {
            $matricula = if ($colMat -gt 0) { $ws.Cells.Item($linha, $colMat).Text.Trim() } else { "" }
            if ([string]::IsNullOrWhiteSpace($matricula)) { continue }

            $nome = if ($colNome -gt 0) { $ws.Cells.Item($linha, $colNome).Text.Trim() } else { "" }
            $email = if ($colEmail -gt 0) { $ws.Cells.Item($linha, $colEmail).Text.Trim() } else { "" }
            if ([string]::IsNullOrWhiteSpace($email)) {
                $email = $formatoEmail.Replace("{matricula}", $matricula).Replace("{nome}", $nome)
            }

            $celMsg1 = if ($colMsg1 -gt 0) { $ws.Cells.Item($linha, $colMsg1) } else { $null }
            $msg1 = if ($celMsg1) {
                if ($celMsg1.Value2) { [string]$celMsg1.Value2.ToString().Trim() } else { [string]$celMsg1.Text.Trim() }
            } else { "" }

            $celMsg2 = if ($colMsg2 -gt 0) { $ws.Cells.Item($linha, $colMsg2) } else { $null }
            $msg2 = if ($celMsg2) {
                if ($celMsg2.Value2) { [string]$celMsg2.Value2.ToString().Trim() } else { [string]$celMsg2.Text.Trim() }
            } else { "" }

            $status = if ($colStatus -gt 0) { $ws.Cells.Item($linha, $colStatus).Text.Trim() } else { "" }
            $numTxt = if ($colNum -gt 0) { $ws.Cells.Item($linha, $colNum).Text.Trim() } else { "0" }
            $num = 0
            [int]::TryParse($numTxt, [ref]$num) | Out-Null

            $dataTxt = if ($colData -gt 0) { $ws.Cells.Item($linha, $colData).Text.Trim() } else { "" }

            # Avaliacao da Elegibilidade com base nas Regras do Configurador
            $apto = $false
            $situacao = ""
            $tipoMsg = if ($num -eq 0) { "1a Msg" } else { "2a Msg (Follow-up)" }
            $mensagemFinal = if ($num -eq 0) { $msg1 } else { if ($msg2) { $msg2 } else { $msg1 } }

            if (-not [string]::IsNullOrWhiteSpace($status)) {
                $situacao = "[Bloqueado] Status: $status"
            }
            elseif ($num -ge $maxContatos) {
                $situacao = "[Bloqueado] Max. atingido ($num/$maxContatos)"
            }
            elseif ($num -eq 0) {
                $apto = $true
                $situacao = "[Apto] 1o Contato"
            }
            else {
                # Segundo contato: verificar se passaram os dias
                $dataAnt = [datetime]::MinValue
                if ([datetime]::TryParse($dataTxt, [ref]$dataAnt)) {
                    $proxima = $dataAnt.AddDays($diasSegundo)
                    if ((Get-Date) -ge $proxima) {
                        $apto = $true
                        $situacao = "[Apto] 2a Msg liberada"
                    }
                    else {
                        $situacao = "[Aguardar] ate " + $proxima.ToString("dd/MM/yyyy")
                    }
                }
                else {
                    $situacao = "[Aviso] Data anterior invalida"
                }
            }

            $item = [PSCustomObject]@{
                Linha        = $linha
                Matricula    = $matricula
                Nome         = $nome
                Destinatario = $email
                Numero       = $num
                TipoMensagem = $tipoMsg
                Situacao     = $situacao
                Status       = $status
                UltimaData   = $dataTxt
                Mensagem1    = $msg1
                Mensagem2    = $msg2
                Mensagem     = $mensagemFinal
                Apto         = $apto
            }
            $global:ListaContatos.Add($item) | Out-Null
        }

        $wb.Close($false)

        $GridPrevia.ItemsSource = $global:ListaContatos
        $aptos = ($global:ListaContatos | Where-Object { $_.Apto }).Count
        $TxtContadores.Text = " Total: $($global:ListaContatos.Count) registros | Aptos para Disparo: $aptos (Limite max: $maxContatos | Espera 2a msg: $diasSegundo d)"
        $TxtStatusGeral.Text = "Previa atualizada: $aptos contato(s) apto(s) para disparo."
        Log-Msg "Regras aplicadas: Max Contatos=$maxContatos, Espera 2a Msg=$diasSegundo dias. Aptos para envio: $aptos."
    }
    catch {
        Log-Msg "Erro ao montar previa: $($_.Exception.Message)"
    }
    finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
    }
}

# Funcao para Carregar e Mapear Colunas da Aba selecionada
function Carregar-Colunas-Mapeamento($salvos = $null) {
    $caminho = $TxtCaminhoPlanilha.Text
    $aba = $CmbAbas.Text

    if (-not (Test-Path $caminho)) {
        Log-Msg "Planilha nao encontrada ou invalida: $caminho"
        return
    }

    Log-Msg "Lendo cabecalhos da aba '$aba'...", $true
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open($caminho, $null, $true)
        $ws = if ($aba) { $wb.Sheets.Item($aba) } else { $wb.Sheets.Item(1) }
        $totalCols = $ws.UsedRange.Columns.Count

        $global:ColunasDetectadas = @()
        for ($c = 1; $c -le $totalCols; $c++) {
            $nomeCol = $ws.Cells.Item(1, $c).Text
            if (-not [string]::IsNullOrWhiteSpace($nomeCol)) {
                $global:ColunasDetectadas += [PSCustomObject]@{
                    Numero = $c
                    Nome   = $nomeCol.Trim()
                }
            }
        }

        $wb.Close($false)

        Log-Msg "Detectadas $($global:ColunasDetectadas.Count) colunas validas no cabecalho."

        # Extrair valores salvos se disponiveis
        $valSalvoMat = if ($salvos) { $salvos.ColMatricula } else { $null }
        $valSalvoNome = if ($salvos) { $salvos.ColNome } else { $null }
        $valSalvoMsg1 = if ($salvos) { $salvos.ColMensagem1 } else { $null }
        $valSalvoMsg2 = if ($salvos) { $salvos.ColMensagem2 } else { $null }
        $valSalvoStatus = if ($salvos) { $salvos.ColStatus } else { $null }
        $valSalvoNum = if ($salvos) { $salvos.ColNumero } else { $null }
        $valSalvoData = if ($salvos) { $salvos.ColUltimaData } else { $null }
        $valSalvoEmail = if ($salvos) { $salvos.ColEmail } else { $null }
        $valSalvoAnalista = if ($salvos) { $salvos.ColAnalista } else { $null }

        Popular-Combo-Colunas $CmbColMatricula $global:ColunasDetectadas @("matricula", "id", "chapa") $true $valSalvoMat
        Popular-Combo-Colunas $CmbColNome $global:ColunasDetectadas @("usuario", "nome") $true $valSalvoNome
        Popular-Combo-Colunas $CmbColMensagem1 $global:ColunasDetectadas @("texto de contato", "mensagem", "contato") $true $valSalvoMsg1
        Popular-Combo-Colunas $CmbColMensagem2 $global:ColunasDetectadas @("texto segundo contato", "segundo contato", "mensagem 2") $false $valSalvoMsg2
        Popular-Combo-Colunas $CmbColStatus $global:ColunasDetectadas @("status envio", "status") $true $valSalvoStatus
        Popular-Combo-Colunas $CmbColNumero $global:ColunasDetectadas @("numero de contato", "numero contato", "contatos") $true $valSalvoNum
        Popular-Combo-Colunas $CmbColUltimaData $global:ColunasDetectadas @("ultima data de contato", "ultima data", "data contato") $true $valSalvoData
        Popular-Combo-Colunas $CmbColEmail $global:ColunasDetectadas @("email_teams", "email", "upn", "matricula") $false $valSalvoEmail
        Popular-Combo-Colunas $CmbColAnalista $global:ColunasDetectadas @("analista", "responsavel", "operador", "enviado por", "usuario envio", "autor", "quem enviou") $false $valSalvoAnalista

        if ($salvos -and -not [string]::IsNullOrWhiteSpace($salvos.NomeAnalista)) {
            $TxtNomeAnalista.Text = $salvos.NomeAnalista
        }

        Log-Msg "Mapeamento preenchido. Carregando previa dos dados..."
        Carregar-Previa-Dados
        Salvar-Configuracoes
    }
    catch {
        Log-Msg "Erro ao ler colunas: $($_.Exception.Message)"
    }
    finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
    }
}

# Carregar Configuracoes Salvas no Inicio
function Carregar-Configuracoes {
    $cfgPath = Join-Path $global:PastaApp "config.json"
    if (-not (Test-Path $cfgPath)) {
        Log-Msg "Nenhum config.json anterior encontrado. Pronto para nova configuracao."
        return
    }

    try {
        $json = [System.IO.File]::ReadAllText($cfgPath, [System.Text.Encoding]::UTF8)
        $cfg = $json | ConvertFrom-Json
        if (-not $cfg) { return }

        Log-Msg "Carregando configuracoes salvas do config.json...", $true

        if ($cfg.MaxContatos) { $TxtMaxContatos.Text = $cfg.MaxContatos }
        if ($cfg.DiasSegundo) { $TxtDiasSegundo.Text = $cfg.DiasSegundo }
        if ($cfg.FormatoEmail) { $TxtFormatoEmail.Text = $cfg.FormatoEmail }
        if ($cfg.EsperaTeams) { $TxtEsperaTeams.Text = $cfg.EsperaTeams }
        if ($cfg.IntervaloEnvio) { $TxtIntervaloEnvio.Text = $cfg.IntervaloEnvio }
        if ($cfg.NomeAnalista) { $TxtNomeAnalista.Text = $cfg.NomeAnalista }
        if ($null -ne $cfg.ConfirmarEnvio) { $ChkConfirmarEnvio.IsChecked = [bool]$cfg.ConfirmarEnvio }
        if ($null -ne $cfg.ModoSimulacao) { $ChkModoSimulacao.IsChecked = [bool]$cfg.ModoSimulacao }

        if (-not [string]::IsNullOrWhiteSpace($cfg.CaminhoPlanilha) -and (Test-Path $cfg.CaminhoPlanilha)) {
            $TxtCaminhoPlanilha.Text = $cfg.CaminhoPlanilha
            Carregar-Abas-Planilha $cfg.CaminhoPlanilha $cfg.AbaSelecionada
            Carregar-Colunas-Mapeamento $cfg
            Log-Msg "Configuracoes salvas restauradas com sucesso!"
        }
    }
    catch {
        Log-Msg "Aviso ao carregar config.json: $($_.Exception.Message)"
    }
}

# Funcao para Exibir a Janela Modal de Resumo dos Envios
function Exibir-Janela-Resumo {
    if ($global:RelatorioFinal.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Nenhum envio registrado para exibir resumo.", "Aviso", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    $readerResumo = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlResumo.OuterXml))
    $modal = [System.Windows.Markup.XamlReader]::Load($readerResumo)
    $modal.Owner = $window

    $TxtResumoTotal = $modal.FindName("TxtResumoTotal")
    $TxtResumoSucessos = $modal.FindName("TxtResumoSucessos")
    $TxtResumoErros = $modal.FindName("TxtResumoErros")
    $GridResumo = $modal.FindName("GridResumo")
    $BtnExportarCsv = $modal.FindName("BtnExportarCsv")
    $BtnFecharResumo = $modal.FindName("BtnFecharResumo")
    $TxtInfoResumo = $modal.FindName("TxtInfoResumo")

    $total = $global:RelatorioFinal.Count
    $sucessos = @($global:RelatorioFinal | Where-Object { $_.Resultado -eq "Sucesso" }).Count
    $erros = @($global:RelatorioFinal | Where-Object { $_.Resultado -ne "Sucesso" }).Count

    $TxtResumoTotal.Text = "$total"
    $TxtResumoSucessos.Text = "$sucessos"
    $TxtResumoErros.Text = "$erros"
    $GridResumo.ItemsSource = $global:RelatorioFinal
    $TxtInfoResumo.Text = "Relatorio com $total itens processados ($sucessos com sucesso, $erros com falha)."

    # Exportar para CSV com delimitador compativel com Excel no Brasil (ponto e virgula)
    $BtnExportarCsv.Add_Click({
        $sfd = New-Object Microsoft.Win32.SaveFileDialog
        $sfd.Filter = "Arquivo CSV (*.csv)|*.csv|Texto (*.txt)|*.txt"
        $sfd.FileName = "Relatorio_Envios_Teams_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv"
        if ($sfd.ShowDialog() -eq $true) {
            try {
                $global:RelatorioFinal | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                [System.Windows.MessageBox]::Show("Relatorio exportado com sucesso para:`n$($sfd.FileName)", "Sucesso", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
            catch {
                [System.Windows.MessageBox]::Show("Erro ao exportar CSV: $($_.Exception.Message)", "Erro", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })

    $BtnFecharResumo.Add_Click({
        $modal.Close()
    })

    $modal.ShowDialog() | Out-Null
}

# Processo de Envio (Individual ou em Massa)
function Executar-Envio($somentePrimeiro = $false) {
    if ($global:ListaContatos.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Carregue e mapeie a planilha primeiro.", "Aviso", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    # Verificar se ainda esta com @empresa.com (exemplo)
    if ($TxtFormatoEmail.Text -like "*@empresa.com*") {
        $aviso = [System.Windows.MessageBox]::Show(
            "Atencao: O campo de e-mail ainda esta com '@empresa.com' (dominio de exemplo).`n`n" +
            "Se voce mantiver @empresa.com, o Teams nao encontrara os usuarios.`n`n" +
            "Deseja ajustar o formato de e-mail agora antes de disparar?",
            "Aviso de Formato do Teams",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($aviso -eq [System.Windows.MessageBoxResult]::Yes) {
            $TxtFormatoEmail.Focus()
            return
        }
    }

    $caminho = $TxtCaminhoPlanilha.Text
    $aba = $CmbAbas.Text
    $colMsg1 = Obter-Indice-Coluna $CmbColMensagem1.Text
    $colMsg2 = Obter-Indice-Coluna $CmbColMensagem2.Text
    $colStatus = Obter-Indice-Coluna $CmbColStatus.Text
    $colNum = Obter-Indice-Coluna $CmbColNumero.Text
    $colData = Obter-Indice-Coluna $CmbColUltimaData.Text
    $colAnalista = Obter-Indice-Coluna $CmbColAnalista.Text
    $nomeAnalista = if ($TxtNomeAnalista -and -not [string]::IsNullOrWhiteSpace($TxtNomeAnalista.Text)) { $TxtNomeAnalista.Text.Trim() } else { Obter-Nome-Analista-Padrao }
    $esperaTeams = [int]$TxtEsperaTeams.Text
    $intervalo = [int]$TxtIntervaloEnvio.Text
    $simulacao = [bool]$ChkModoSimulacao.IsChecked
    $pedirConfirmacao = [bool]$ChkConfirmarEnvio.IsChecked

    $aptos = @($global:ListaContatos | Where-Object { $_.Apto })
    if ($aptos.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Nao ha contatos aptos para envio no momento.", "Informacao", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    if ($somentePrimeiro) {
        $aptos = @($aptos[0])
        Log-Msg "Modo Teste: Enviando apenas para o primeiro contato ($($aptos[0].Matricula)) por '$nomeAnalista'...", $true
    }
    else {
        $resp = [System.Windows.MessageBox]::Show("Confirmar inicio do disparo automatico para $($aptos.Count) contatos aptos?`n`nAnalista responsavel: $nomeAnalista", "Confirmacao de Envio", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($resp -ne [System.Windows.MessageBoxResult]::Yes) { return }
        Log-Msg "Iniciando disparo em massa para $($aptos.Count) contatos por '$nomeAnalista'...", $true
    }

    # Criar backup antes de iniciar
    if (-not $simulacao) {
        $dirName = [System.IO.Path]::GetDirectoryName($caminho)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($caminho)
        $ext = [System.IO.Path]::GetExtension($caminho)
        $timestampBackup = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $backupPath = [System.IO.Path]::Combine($dirName, "${baseName}_backup_${timestampBackup}${ext}")
        Copy-Item -Path $caminho -Destination $backupPath -Force
        Log-Msg "Backup de seguranca criado: $(Split-Path $backupPath -Leaf)"
    }

    $BtnIniciar.IsEnabled = $false
    $BtnTestarUm.IsEnabled = $false
    $BtnParar.IsEnabled = $true
    $global:PararSolicitado = $false
    $BarraProgresso.Maximum = $aptos.Count
    $BarraProgresso.Value = 0

    # Inicializar a lista do relatorio de resumo
    $global:RelatorioFinal = [System.Collections.ArrayList]::new()

    $excel = $null
    $wb = $null
    $ws = $null

    $enviados = 0
    $falhas = 0

    try {
        if (-not $simulacao) {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $wb = $excel.Workbooks.Open($caminho)
            $excel.Calculation = -4105 # xlCalculationAutomatic
            try { $wb.ForceFullCalculation = $true } catch {}
            $excel.Calculate()
            $excel.CalculateFull()
            $ws = if ($aba) { $wb.Sheets.Item($aba) } else { $wb.Sheets.Item(1) }
        }

        for ($i = 0; $i -lt $aptos.Count; $i++) {
            $c = $aptos[$i]

            if ($global:PararSolicitado) {
                Log-Msg "Envio interrompido pelo operador."
                for ($j = $i; $j -lt $aptos.Count; $j++) {
                    $itemPendente = $aptos[$j]
                    $global:RelatorioFinal.Add([PSCustomObject]@{
                        Linha        = $itemPendente.Linha
                        Matricula    = $itemPendente.Matricula
                        Nome         = $itemPendente.Nome
                        Destinatario = $itemPendente.Destinatario
                        Analista     = $nomeAnalista
                        Tipo         = $itemPendente.TipoMensagem
                        Resultado    = "Nao Enviado"
                        Motivo       = "Interrompido pelo operador antes do envio (Excel nao alterado)"
                    }) | Out-Null
                    $falhas++
                }
                break
            }

            $TxtStatusGeral.Text = "Processando $($i+1) de $($aptos.Count): $($c.Nome) ($($c.Matricula))..."
            Log-Msg "[$($i+1)/$($aptos.Count)] Destinatario: $($c.Destinatario) | Tipo: $($c.TipoMensagem) | Matricula: $($c.Matricula) | Analista: $nomeAnalista"

            # Validar destinatario e mensagem antes de disparar
            if ([string]::IsNullOrWhiteSpace($c.Destinatario)) {
                $erroMsg = "Destinatario vazio na planilha"
                Log-Msg "  [FALHA] $erroMsg - Excel NAO alterado."
                $c.Situacao = "[Falha] $erroMsg"
                $global:RelatorioFinal.Add([PSCustomObject]@{
                    Linha        = $c.Linha
                    Matricula    = $c.Matricula
                    Nome         = $c.Nome
                    Destinatario = $c.Destinatario
                    Analista     = $nomeAnalista
                    Tipo         = $c.TipoMensagem
                    Resultado    = "Erro"
                    Motivo       = "$erroMsg (Excel preservado)"
                }) | Out-Null
                $falhas++
                $BarraProgresso.Value = $i + 1
                continue
            }

            # Recalcular e puxar o texto da mensagem mais recente direto da celula do Excel
            if (-not $simulacao -and $ws) {
                $excel.Calculate()
                $celMsg1Atual = if ($colMsg1 -gt 0) { $ws.Cells.Item($c.Linha, $colMsg1) } else { $null }
                if ($celMsg1Atual) {
                    $txtMsg1 = if ($celMsg1Atual.Value2) { [string]$celMsg1Atual.Value2.ToString().Trim() } else { [string]$celMsg1Atual.Text.Trim() }
                    if (-not [string]::IsNullOrWhiteSpace($txtMsg1)) { $c.Mensagem1 = $txtMsg1 }
                }

                $celMsg2Atual = if ($colMsg2 -gt 0) { $ws.Cells.Item($c.Linha, $colMsg2) } else { $null }
                if ($celMsg2Atual) {
                    $txtMsg2 = if ($celMsg2Atual.Value2) { [string]$celMsg2Atual.Value2.ToString().Trim() } else { [string]$celMsg2Atual.Text.Trim() }
                    if (-not [string]::IsNullOrWhiteSpace($txtMsg2)) { $c.Mensagem2 = $txtMsg2 }
                }

                $celulaMsgAtual = if ($c.Numero -eq 0) { $celMsg1Atual } else { if ($celMsg2Atual) { $celMsg2Atual } else { $celMsg1Atual } }
                if ($celulaMsgAtual) {
                    $txtAtual = if ($celulaMsgAtual.Value2) { [string]$celulaMsgAtual.Value2.ToString().Trim() } else { [string]$celulaMsgAtual.Text.Trim() }
                    if (-not [string]::IsNullOrWhiteSpace($txtAtual)) {
                        $c.Mensagem = $txtAtual
                    }
                }
            }

            if ([string]::IsNullOrWhiteSpace($c.Mensagem)) {
                $erroMsg = "Texto da mensagem vazia na planilha"
                Log-Msg "  [FALHA] $erroMsg - Excel NAO alterado."
                $c.Situacao = "[Falha] $erroMsg"
                $global:RelatorioFinal.Add([PSCustomObject]@{
                    Linha        = $c.Linha
                    Matricula    = $c.Matricula
                    Nome         = $c.Nome
                    Destinatario = $c.Destinatario
                    Analista     = $nomeAnalista
                    Tipo         = $c.TipoMensagem
                    Resultado    = "Erro"
                    Motivo       = "$erroMsg (Excel preservado)"
                }) | Out-Null
                $falhas++
                $BarraProgresso.Value = $i + 1
                continue
            }

            $sucessoItem = $false
            try {
                if ($simulacao) {
                    Log-Msg "  [SIMULACAO] Testando envio sem disparar mensagem real..."
                    Start-Sleep -Seconds 1
                    $sucessoItem = $true
                }
                else {
                    # 1. Abrir ou focar o Teams nativo
                    # O protocolo msteams:/l/chat/0/0 abre a tela de chat do Teams diretamente
                    Start-Process "msteams:/l/chat/0/0"
                    Start-Sleep -Seconds $esperaTeams

                    # 2. Localizar janela real do Teams, focar e maximizar (funciona 100% no Novo Teams WebView2)
                    try {
                        [TeamsHelper]::AtivarEMaximizar([IntPtr]::Zero)
                        Start-Sleep -Milliseconds 400
                    }
                    catch { }

                    # 3. Fechar qualquer menu ou popup aberto com ESC
                    [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
                    Start-Sleep -Milliseconds 200

                    # 4. ACESSAR A BARRA DE PESQUISA: Clique fisico centralizado na barra Pesquisar + atalho Ctrl+E
                    # O clique na barra transfere o foco fisico do Windows para o Teams e ativa o cursor
                    try {
                        [TeamsHelper]::ClicarNaBarraPesquisa()
                        Start-Sleep -Milliseconds 300
                    } catch { }
                    [System.Windows.Forms.SendKeys]::SendWait("^e")
                    Start-Sleep -Milliseconds 400

                    # 5. Limpar qualquer busca anterior na barra com Backspace seguro
                    [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}{BACKSPACE}")
                    Start-Sleep -Milliseconds 150

                    # 6. Digitar a matricula caractere por caractere diretamente na pesquisa
                    $destLimpo = $c.Destinatario.Trim()
                    Log-Msg "  Pesquisando destinatario '$destLimpo' via Pesquisa (Ctrl+E)..."
                    [System.Windows.Forms.SendKeys]::SendWait($destLimpo)
                    Start-Sleep -Milliseconds 400

                    # Toque de espaco e backspace para disparar o autocomplete do Active Directory / Entra ID do Santander
                    [System.Windows.Forms.SendKeys]::SendWait(" {BACKSPACE}")

                    # 7. Aguardar o Teams buscar e sugerir o colaborador no dropdown do Santander
                    Start-Sleep -Milliseconds 2800

                    # 8. Selecionar o colaborador sugerido na pesquisa e abrir o chat direto
                    # O Novo Teams exibe uma linha de filtros (from:...) antes do colaborador.
                    # Pressionamos 2 vezes a seta para baixo ({DOWN}) para selecionar o colaborador e Enter para abrir o chat.
                    [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
                    Start-Sleep -Milliseconds 250
                    [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
                    Start-Sleep -Milliseconds 300
                    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                    Start-Sleep -Milliseconds 2000 # Aguarda o Teams carregar o chat com o colaborador

                    # 9. Ir do campo 'Para:' para a caixa de mensagem 'Digite uma mensagem'
                    # No Teams, ao confirmar o destinatario, pressionar ENTER ou TAB pula direto para a mensagem
                    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                    Start-Sleep -Milliseconds 250
                    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
                    Start-Sleep -Milliseconds 250

                    # Atalho oficial do Teams para ir direto para a caixa de texto: Ctrl + R
                    [System.Windows.Forms.SendKeys]::SendWait("^r")
                    Start-Sleep -Milliseconds 250

                    # Clique fisico assistido calibrado exatamente dentro da caixa de texto (75px acima do rodape)
                    try {
                        [TeamsHelper]::ClicarNoCampoMensagem()
                        Start-Sleep -Milliseconds 400
                    } catch { }

                    # 10. Copiar e colar a mensagem personalizada atualizada
                    Log-Msg "  Colando mensagem personalizada..."
                    [System.Windows.Forms.Clipboard]::SetText($c.Mensagem)
                    Start-Sleep -Milliseconds 200
                    [System.Windows.Forms.SendKeys]::SendWait("^v")
                    Start-Sleep -Milliseconds 800

                    # 11. DISPARAR O ENVIO DA MENSAGEM NO TEAMS
                    # Enviamos Ctrl+Enter (atalho universal de envio da Microsoft), Enter e clique fisico no botao Enviar
                    Log-Msg "  Disparando envio (Ctrl+Enter e Botao Enviar)..."
                    [System.Windows.Forms.SendKeys]::SendWait("^{ENTER}")
                    Start-Sleep -Milliseconds 300
                    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                    Start-Sleep -Milliseconds 300
                    try {
                        [TeamsHelper]::ClicarNoBotaoEnviar()
                        Start-Sleep -Milliseconds 500
                    } catch { }

                    # Validacao com operador SOMENTE se a opcao estiver explicitamente marcada na tela
                    if ($pedirConfirmacao) {
                        $respostaOperador = [System.Windows.MessageBox]::Show(
                            "A mensagem foi enviada para $($c.Nome) ($($c.Destinatario))?`n`n" +
                            "Clique em 'SIM' para confirmar a gravacao no Excel.`n" +
                            "Clique em 'NAO' para cancelar a gravacao deste contato.",
                            "Validacao de Envio",
                            [System.Windows.MessageBoxButton]::YesNo,
                            [System.Windows.MessageBoxImage]::Question
                         )
                        if ($respostaOperador -eq [System.Windows.MessageBoxResult]::Yes) {
                            $sucessoItem = $true
                        }
                        else {
                            $sucessoItem = $false
                            Log-Msg "  [AVISO] Envio cancelado pelo operador na confirmacao."
                        }
                    }
                    else {
                        # Modo 100% automatico sem interrupcao de popup
                        $sucessoItem = $true
                    }

                    # ATUALIZAR EXCEL SOMENTE SE HOUVE SUCESSO CONFIRMADO!
                    if ($sucessoItem) {
                        try {
                            # 1. Atualizar contagem de contatos
                            if ($colNum -gt 0) {
                                $novoNum = [int]$c.Numero + 1
                                try {
                                    $ws.Cells.Item([int]$c.Linha, [int]$colNum).Value = $novoNum
                                } catch {
                                    $ws.Cells.Item([int]$c.Linha, [int]$colNum).Value2 = [double]$novoNum
                                }
                            }

                            # 2. Atualizar data do ultimo contato (com formatacao texto segura para evitar InvalidCastException)
                            if ($colData -gt 0) {
                                $dataHoraStr = (Get-Date).ToString("dd/MM/yyyy HH:mm")
                                try {
                                    $ws.Cells.Item([int]$c.Linha, [int]$colData).NumberFormat = "@"
                                    $ws.Cells.Item([int]$c.Linha, [int]$colData).Value = $dataHoraStr
                                } catch {
                                    try {
                                        $ws.Cells.Item([int]$c.Linha, [int]$colData).Value2 = $dataHoraStr
                                    } catch {
                                        $ws.Cells.Item([int]$c.Linha, [int]$colData) = $dataHoraStr
                                    }
                                }
                            }

                            # 3. Atualizar status do envio na planilha se coluna estiver configurada
                            if ($colStatus -gt 0) {
                                try {
                                    $ws.Cells.Item([int]$c.Linha, [int]$colStatus).Value = "Enviado"
                                } catch {
                                    $ws.Cells.Item([int]$c.Linha, [int]$colStatus).Value2 = "Enviado"
                                }
                            }

                            # 4. Atualizar nome do analista na planilha se coluna estiver configurada
                            if ($colAnalista -gt 0 -and -not [string]::IsNullOrWhiteSpace($nomeAnalista)) {
                                try {
                                    $ws.Cells.Item([int]$c.Linha, [int]$colAnalista).NumberFormat = "@"
                                    $ws.Cells.Item([int]$c.Linha, [int]$colAnalista).Value = $nomeAnalista
                                } catch {
                                    try {
                                        $ws.Cells.Item([int]$c.Linha, [int]$colAnalista).Value2 = $nomeAnalista
                                    } catch {
                                        $ws.Cells.Item([int]$c.Linha, [int]$colAnalista) = $nomeAnalista
                                    }
                                }
                            }

                            $wb.Save()
                            Log-Msg "  -> Mensagem enviada por '$nomeAnalista' e planilha atualizada com sucesso!"
                        }
                        catch {
                            Log-Msg "  -> Mensagem enviada no Teams, mas houve aviso ao salvar no Excel: $($_.Exception.Message)"
                        }

                        $c.Situacao = "[Sucesso] Enviado"
                        $c.Apto = $false
                        $enviados++

                        $global:RelatorioFinal.Add([PSCustomObject]@{
                            Linha        = $c.Linha
                            Matricula    = $c.Matricula
                            Nome         = $c.Nome
                            Destinatario = $c.Destinatario
                            Analista     = $nomeAnalista
                            Tipo         = $c.TipoMensagem
                            Resultado    = "Sucesso"
                            Motivo       = "Enviado com sucesso no Teams e confirmado no Excel"
                        }) | Out-Null
                    }
                    else {
                        $falhas++
                        $c.Situacao = "[Falha] Nao confirmado / Erro no Teams"
                        $global:RelatorioFinal.Add([PSCustomObject]@{
                            Linha        = $c.Linha
                            Matricula    = $c.Matricula
                            Nome         = $c.Nome
                            Destinatario = $c.Destinatario
                            Analista     = $nomeAnalista
                            Tipo         = $c.TipoMensagem
                            Resultado    = "Erro"
                            Motivo       = "Envio nao confirmado no Teams (Excel NAO alterado)"
                        }) | Out-Null
                    }
                }
            }
            catch {
                $erroEx = "Falha durante o envio: $($_.Exception.Message)"
                Log-Msg "  [ERRO] $erroEx (Excel NAO alterado)"
                $c.Situacao = "[Falha] $($_.Exception.Message)"
                $falhas++

                $global:RelatorioFinal.Add([PSCustomObject]@{
                    Linha        = $c.Linha
                    Matricula    = $c.Matricula
                    Nome         = $c.Nome
                    Destinatario = $c.Destinatario
                    Analista     = $nomeAnalista
                    Tipo         = $c.TipoMensagem
                    Resultado    = "Erro"
                    Motivo       = "$erroEx (Excel preservado)"
                }) | Out-Null
            }

            $BarraProgresso.Value = $i + 1
            [System.Windows.Forms.Application]::DoEvents()

            # Intervalo entre envios (se nao for o ultimo)
            if ($i -lt ($aptos.Count - 1)) {
                Start-Sleep -Seconds $intervalo
            }
        }
    }
    catch {
        Log-Msg "Erro geral durante o processamento: $($_.Exception.Message)"
    }
    finally {
        if ($wb) {
            $wb.Save()
            $wb.Close($false)
        }
        if ($excel) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect()
        }

        $BtnIniciar.IsEnabled = $true
        $BtnTestarUm.IsEnabled = $true
        $BtnParar.IsEnabled = $false
        $BtnVerUltimoResumo.IsEnabled = $true
        $TxtStatusGeral.Text = "Finalizado: $enviados enviado(s), $falhas falha(s)."
        Log-Msg "Processamento concluido: $enviados enviado(s) com sucesso, $falhas com erro/falha.", $true
        $GridPrevia.Items.Refresh()

        # Abrir automaticamente a Tela Modal de Resumo dos Envios
        Exibir-Janela-Resumo
    }
}

# Botoes de Configuracao e Recalculo
$BtnCarregarColunas.Add_Click({
    Carregar-Colunas-Mapeamento $null
})

$BtnRecalcular.Add_Click({
    Log-Msg "Recalculando previa com novas regras de limite e 2a mensagem..."
    Carregar-Previa-Dados
    Salvar-Configuracoes
})

$BtnSalvarConfig.Add_Click({
    Salvar-Configuracoes
    [System.Windows.MessageBox]::Show("Configurações salvas com sucesso em config.json!`n`nNa próxima vez que o aplicativo for aberto, a planilha, aba, colunas e regras serão restauradas automaticamente.", "Configurações Salvas", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
})

$BtnVerUltimoResumo.Add_Click({
    Exibir-Janela-Resumo
})

# Botoes de Disparo
$BtnIniciar.Add_Click({ Executar-Envio $false })
$BtnTestarUm.Add_Click({ Executar-Envio $true })
$BtnParar.Add_Click({
    $global:PararSolicitado = $true
    Log-Msg "Solicitacao de parada enviada. Aguardando conclusao do item atual..."
})

# Eventos de Inicializacao e Salvamento Automatico
$window.Add_ContentRendered({
    Carregar-Configuracoes
})

$window.Add_Closing({
    Salvar-Configuracoes
})

# Exibir a Janela Principal
$window.ShowDialog() | Out-Null
