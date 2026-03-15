[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$script:RootPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:EngineModulePath = Join-Path $script:RootPath 'Modules\AzureSupport.TicketEngine.psm1'
$script:RequestRows = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
$script:DiscoveryJob = $null
$script:RunJob = $null
$script:RegionJob = $null
$script:PollTimer = $null
$script:CancellationFile = $null
$script:RunId = $null
$script:RunStatePath = $null
$script:ProfilePath = Join-Path $script:RootPath 'config\azure-ticket-gui-profile.json'
$script:TicketTemplatePath = Join-Path $script:RootPath 'config\default-ticket-template.json'
$script:Template = $null
$script:ValidationState = @{ IsValid = $false; Errors = @(); Warnings = @() }
$script:LogFilePath = $null
$script:AvailableRegions = @()

$script:LogDir = Join-Path $script:RootPath 'logs'
if (-not (Test-Path -LiteralPath $script:LogDir)) {
    New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
}
$script:LogFilePath = Join-Path $script:LogDir ("gui-" + (Get-Date -Format 'yyyy-MM-dd') + ".log")

if (-not (Test-Path -LiteralPath $script:EngineModulePath)) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "Ticket engine module not found at '$($script:EngineModulePath)'.`nThe GUI cannot start without it.",
        "Azure Support Ticket Engine",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error)
    exit 1
}
try {
    Import-Module -Name $script:EngineModulePath -Force -ErrorAction Stop | Out-Null
}
catch {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "Unable to import ticket engine module: $($_.Exception.Message)",
        "Azure Support Ticket Engine",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error)
    exit 1
}

function Load-TicketTemplate {
    param([Parameter(Mandatory = $false)][string]$TicketTemplatePath)
    $resolved = if ([string]::IsNullOrWhiteSpace($TicketTemplatePath)) {
        Get-TicketTemplatePath -RootPath $script:RootPath
    } else { $TicketTemplatePath }
    return Get-TicketTemplate -Path $resolved
}

try {
    $script:Template = Load-TicketTemplate -TicketTemplatePath $script:TicketTemplatePath
}
catch {
    $fallbackTemplate = [pscustomobject]@{
        title = 'Quota request for Batch'
        acceptLanguage = 'en'
        defaults = @{ delaySeconds = 23; requestsPerMinute = 2; maxRetries = 6; baseRetrySeconds = 25; newLimit = 680; quotaType = 'LowPriority' }
        contactDetails = @{
            firstName = 'Support'; lastName = 'User'; preferredContactMethod = 'email'
            primaryEmailAddress = 'support@example.com'; preferredTimeZone = 'UTC'
            country = 'US'; preferredSupportLanguage = 'en-us'; additionalEmailAddresses = @()
        }
        problemClassificationId = '/providers/microsoft.support/services/06bfd9d3-516b-d5c6-5802-169c800dec89/problemclassifications/831b2fb3-4db3-3d32-af35-bbb3d3eaeba2'
        serviceId = '/providers/microsoft.support/services/06bfd9d3-516b-d5c6-5802-169c800dec89'
        severity = 'minimal'
        descriptionTemplate = 'Quota request for Batch'
        advancedDiagnosticConsent = 'Yes'
        require24X7Response = $false
        supportPlanId = 'U291cmNlOkZyZWUsRnJlZUlkOjAwMDAwMDAwLTAwMDAtMDAwMC0wMDAwLTAwMDAwMDAwMDAwOSw%3d'
        quotaChangeRequestVersion = '1.0'
        quotaChangeRequestSubType = 'Account'
        quotaRequestType = 'LowPriority'
        defaultRequests = @()
    }
    $script:Template = $fallbackTemplate
}

$script:Defaults = Merge-TemplateDefaults -Template $script:Template

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="MainWindow"
    Title="Azure Support Ticket Engine"
    Width="1360"
    Height="940"
    WindowStartupLocation="CenterScreen"
    Background="#F0F2F5"
    FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="TextWrapping" Value="Wrap" />
            <Setter Property="Foreground" Value="#1B1B1B" />
        </Style>
        <Style x:Key="AccentButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4" />
            <Setter Property="Foreground" Value="White" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="14 6" />
            <Setter Property="FontWeight" Value="SemiBold" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#106EBE" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#005A9E" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#A0A0A0" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="#E1E1E1" />
            <Setter Property="Foreground" Value="#1B1B1B" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Padding" Value="14 6" />
            <Setter Property="FontWeight" Value="Normal" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#D0D0D0" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#B8B8B8" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#C8C8C8" />
                                <Setter Property="Foreground" Value="#888" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header Bar -->
        <Border Grid.Row="0" Background="#0078D4" Padding="16 10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="&#x2601; Azure Support Ticket Engine"
                           FontSize="20" FontWeight="Bold" Foreground="White" VerticalAlignment="Center"/>
                <TextBlock x:Name="TxtSummary" Grid.Column="1" Foreground="#D0E8FF"
                           FontSize="13" HorizontalAlignment="Right" VerticalAlignment="Center"/>
            </Grid>
        </Border>

        <!-- Main Content -->
        <Grid Grid.Row="1" Margin="12 8 12 4">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="200"/>
            </Grid.RowDefinitions>

            <!-- Discovery Filters -->
            <Border Grid.Row="0" Background="White" CornerRadius="6" Padding="12" Margin="0 0 0 8"
                    BorderBrush="#D0D8E0" BorderThickness="1">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Row="0" Grid.Column="0" Margin="0 0 8 0">
                        <TextBlock FontWeight="SemiBold" Text="Subscription IDs" FontSize="12" Foreground="#555"/>
                        <TextBox x:Name="TxtSubscriptionFilter" Margin="0 3 0 0" Padding="4"/>
                    </StackPanel>
                    <StackPanel Grid.Row="0" Grid.Column="1" Margin="0 0 8 0">
                        <TextBlock FontWeight="SemiBold" Text="Tenant ID contains" FontSize="12" Foreground="#555"/>
                        <TextBox x:Name="TxtTenantFilter" Margin="0 3 0 0" Padding="4"/>
                    </StackPanel>
                    <StackPanel Grid.Row="0" Grid.Column="2" Margin="0 0 8 0">
                        <TextBlock FontWeight="SemiBold" Text="Region / Account contains" FontSize="12" Foreground="#555"/>
                        <TextBox x:Name="TxtRegionFilter" Margin="0 3 0 0" Padding="4"/>
                        <TextBox x:Name="TxtAccountFilter" Margin="0 4 0 0" Padding="4" ToolTip="Account name contains..."/>
                    </StackPanel>

                    <StackPanel Grid.Row="0" Grid.Column="3" VerticalAlignment="Bottom" Margin="4 0 0 0">
                        <Button x:Name="BtnDiscover" Style="{StaticResource AccentButton}"
                                Width="140" Height="32" Content="Discover Accounts" Margin="0 0 0 4"/>
                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="BtnSelectAll" Style="{StaticResource SecondaryButton}"
                                    Width="66" Height="28" Content="All" Margin="0 0 4 0"/>
                            <Button x:Name="BtnClearAll" Style="{StaticResource SecondaryButton}"
                                    Width="66" Height="28" Content="None"/>
                        </StackPanel>
                    </StackPanel>

                    <StackPanel Grid.Row="1" Grid.ColumnSpan="4" Margin="0 6 0 0">
                        <ProgressBar x:Name="DiscoveryProgress" Height="3" IsIndeterminate="False"
                                     Foreground="#0078D4" Background="#E8E8E8" Margin="0 0 0 2"/>
                        <TextBlock x:Name="TxtStatusHint" Foreground="#6B7280" FontSize="11.5"
                                   Text="Use filters to scope discovery, then select rows before execution."/>
                    </StackPanel>
                </Grid>
            </Border>

            <!-- Middle Section: DataGrid + Settings -->
            <Grid Grid.Row="1" Margin="0 0 0 6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="2.5*"/>
                    <ColumnDefinition Width="2*"/>
                </Grid.ColumnDefinitions>

                <!-- Request Grid -->
                <Border Grid.Column="0" Background="White" CornerRadius="6"
                        BorderBrush="#D0D8E0" BorderThickness="1" Padding="8" Margin="0 0 6 0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" FontSize="13" FontWeight="SemiBold"
                                   Margin="0 0 0 6" Text="Discovered Requests"/>
                        <DataGrid
                            x:Name="RequestGrid"
                            Grid.Row="1"
                            AutoGenerateColumns="False"
                            CanUserAddRows="False"
                            IsReadOnly="False"
                            AlternationCount="2"
                            HeadersVisibility="Column"
                            SelectionMode="Extended"
                            GridLinesVisibility="Horizontal"
                            HorizontalGridLinesBrush="#EEE"
                            RowHeight="28"
                            FontSize="12">
                            <DataGrid.Columns>
                                <DataGridCheckBoxColumn Header="Run" Binding="{Binding Selected}" Width="45"/>
                                <DataGridTextColumn Header="Subscription" Binding="{Binding SubscriptionName}" IsReadOnly="True" Width="1.6*"/>
                                <DataGridTextColumn Header="Subscription ID" Binding="{Binding SubscriptionId}" IsReadOnly="True" Width="1.8*"/>
                                <DataGridTextColumn Header="Account" Binding="{Binding AccountName}" IsReadOnly="True" Width="1.4*"/>
                                <DataGridTextColumn Header="Region" Binding="{Binding Region, Mode=TwoWay}" Width="0.9*"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding Status}" IsReadOnly="True" Width="0.8*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Grid>
                </Border>

                <!-- Settings Panel with Tabs -->
                <Border Grid.Column="1" Background="White" CornerRadius="6"
                        BorderBrush="#D0D8E0" BorderThickness="1" Padding="8">
                    <TabControl x:Name="SettingsTabs" BorderThickness="0" Background="Transparent">
                        <!-- Execution Tab -->
                        <TabItem Header="Execution" Padding="8 4">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0 6 0 0">
                                <StackPanel>
                                    <TextBlock FontWeight="SemiBold" Text="Bearer Token" FontSize="12"/>
                                    <PasswordBox x:Name="TxtToken" Margin="0 2 0 4" Padding="3"/>
                                    <CheckBox x:Name="ChkTryAzCliToken" Content="Use Azure CLI token when empty" IsChecked="True" Margin="0 0 0 4"/>
                                    <CheckBox x:Name="ChkUseDeviceCode" Content="Allow device-code token fallback" IsChecked="False" Margin="0 0 0 8"/>

                                    <TextBlock FontWeight="SemiBold" Text="Timing" FontSize="12"/>
                                    <Grid Margin="0 3 0 0">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="1*"/>
                                            <ColumnDefinition Width="1*"/>
                                            <ColumnDefinition Width="1*"/>
                                        </Grid.ColumnDefinitions>
                                        <StackPanel Grid.Column="0" Margin="0 0 4 0">
                                            <TextBlock Text="Delay (s)" FontSize="11" Foreground="#666"/>
                                            <TextBox x:Name="TxtDelaySeconds" Padding="3"/>
                                        </StackPanel>
                                        <StackPanel Grid.Column="1" Margin="0 0 4 0">
                                            <TextBlock Text="Req/min" FontSize="11" Foreground="#666"/>
                                            <TextBox x:Name="TxtRequestsPerMinute" Padding="3"/>
                                        </StackPanel>
                                        <StackPanel Grid.Column="2">
                                            <TextBlock Text="Retries" FontSize="11" Foreground="#666"/>
                                            <TextBox x:Name="TxtMaxRetries" Padding="3"/>
                                        </StackPanel>
                                    </Grid>
                                    <StackPanel Margin="0 4 0 8">
                                        <TextBlock Text="Retry base (s)" FontSize="11" Foreground="#666"/>
                                        <TextBox x:Name="TxtBaseRetrySeconds" Padding="3" Width="100" HorizontalAlignment="Left"/>
                                    </StackPanel>

                                    <TextBlock FontWeight="SemiBold" Text="Quota" FontSize="12"/>
                                    <Grid Margin="0 3 0 0">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="1*"/>
                                            <ColumnDefinition Width="1*"/>
                                        </Grid.ColumnDefinitions>
                                        <StackPanel Grid.Column="0" Margin="0 0 4 0">
                                            <TextBlock Text="New limit" FontSize="11" Foreground="#666"/>
                                            <TextBox x:Name="TxtNewLimit" Padding="3"/>
                                        </StackPanel>
                                        <StackPanel Grid.Column="1">
                                            <TextBlock Text="Quota type" FontSize="11" Foreground="#666"/>
                                            <ComboBox x:Name="CmbQuotaType" Padding="3">
                                                <ComboBoxItem>LowPriority</ComboBoxItem>
                                                <ComboBoxItem>Dedicated</ComboBoxItem>
                                            </ComboBox>
                                        </StackPanel>
                                    </Grid>
                                    <TextBlock Text="Ticket title" FontSize="11" Foreground="#666" Margin="0 6 0 0"/>
                                    <TextBox x:Name="TxtTitle" Padding="3" Margin="0 2 0 8"/>

                                    <CheckBox x:Name="ChkRotateFingerprint" Content="Rotate fingerprint headers" IsChecked="True" Margin="0 0 0 4"/>
                                    <CheckBox x:Name="ChkDryRun" Content="Dry run (no API calls)" IsChecked="False" Margin="0 0 0 4"/>
                                    <CheckBox x:Name="ChkStopOnFirstFailure" Content="Stop on first failure" IsChecked="False" Margin="0 0 0 4"/>
                                    <CheckBox x:Name="ChkResumeFromState" Content="Resume from previous run" IsChecked="False" Margin="0 0 0 4"/>
                                    <CheckBox x:Name="ChkRetryFailedRequests" Content="Retry failed requests only" IsChecked="False" Margin="0 0 0 8"/>

                                    <TextBlock FontWeight="SemiBold" Text="Proxy (optional)" FontSize="12"/>
                                    <TextBox x:Name="TxtProxy" Padding="3" Margin="0 2 0 4"/>
                                    <CheckBox x:Name="ChkProxyDefault" Content="Use default proxy credentials" IsChecked="False" Margin="0 0 0 8"/>

                                    <TextBlock FontWeight="SemiBold" Text="Artifact Paths" FontSize="12"/>
                                    <TextBlock Text="Run state JSON" FontSize="11" Foreground="#666"/>
                                    <TextBox x:Name="TxtRunStatePath" Padding="3" Margin="0 2 0 4"/>
                                    <TextBlock Text="Result JSON" FontSize="11" Foreground="#666"/>
                                    <TextBox x:Name="TxtResultJsonPath" Padding="3" Margin="0 2 0 4"/>
                                    <TextBlock Text="Result CSV" FontSize="11" Foreground="#666"/>
                                    <TextBox x:Name="TxtResultCsvPath" Padding="3" Margin="0 2 0 4"/>
                                    <TextBlock Text="Ticket template" FontSize="11" Foreground="#666"/>
                                    <TextBox x:Name="TxtTicketTemplatePath" Padding="3" Margin="0 2 0 4"/>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>

                        <!-- Contact Tab -->
                        <TabItem Header="Contact" Padding="8 4">
                            <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0 6 0 0">
                                <StackPanel>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="1*"/>
                                            <ColumnDefinition Width="1*"/>
                                        </Grid.ColumnDefinitions>
                                        <StackPanel Grid.Column="0" Margin="0 0 4 0">
                                            <TextBlock Text="First name" FontSize="12" FontWeight="SemiBold"/>
                                            <TextBox x:Name="TxtContactFirstName" Padding="3" Margin="0 2 0 8"/>
                                        </StackPanel>
                                        <StackPanel Grid.Column="1">
                                            <TextBlock Text="Last name" FontSize="12" FontWeight="SemiBold"/>
                                            <TextBox x:Name="TxtContactLastName" Padding="3" Margin="0 2 0 8"/>
                                        </StackPanel>
                                    </Grid>
                                    <TextBlock Text="Email" FontSize="12" FontWeight="SemiBold"/>
                                    <TextBox x:Name="TxtContactEmail" Padding="3" Margin="0 2 0 8"/>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="1*"/>
                                            <ColumnDefinition Width="1*"/>
                                        </Grid.ColumnDefinitions>
                                        <StackPanel Grid.Column="0" Margin="0 0 4 0">
                                            <TextBlock Text="Time zone" FontSize="12" FontWeight="SemiBold"/>
                                            <TextBox x:Name="TxtTimezone" Padding="3" Margin="0 2 0 8"/>
                                        </StackPanel>
                                        <StackPanel Grid.Column="1">
                                            <TextBlock Text="Support language" FontSize="12" FontWeight="SemiBold"/>
                                            <TextBox x:Name="TxtSupportLanguage" Padding="3" Margin="0 2 0 8"/>
                                        </StackPanel>
                                    </Grid>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>

                        <!-- Regions Tab -->
                        <TabItem Header="Regions" Padding="8 4">
                            <Grid Margin="0 6 0 0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" FontSize="12" FontWeight="SemiBold"
                                           Text="Available Azure Regions (select up to 10)" Margin="0 0 0 4"/>
                                <TextBlock x:Name="TxtRegionCount" Grid.Row="1" FontSize="11"
                                           Foreground="#666" Text="Loading regions..." Margin="0 0 0 4"/>
                                <ListBox x:Name="RegionListBox" Grid.Row="2"
                                         SelectionMode="Multiple"
                                         BorderThickness="1" BorderBrush="#D0D8E0"
                                         ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                            </Grid>
                        </TabItem>
                    </TabControl>
                </Border>
            </Grid>

            <!-- Bottom Section: Validation, Run Controls, Log -->
            <Grid Grid.Row="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Run Controls -->
                <Border Grid.Row="0" Background="White" CornerRadius="6" Padding="10" Margin="0 0 0 6"
                        BorderBrush="#D0D8E0" BorderThickness="1">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name="BtnRun" Grid.Column="0" Style="{StaticResource AccentButton}"
                                Width="130" Height="32" Content="Run Selected"/>
                        <Button x:Name="BtnCancelRun" Grid.Column="1" Style="{StaticResource SecondaryButton}"
                                Width="100" Height="32" Margin="8 0 0 0" Content="Cancel"/>
                        <Button x:Name="BtnReloadDefaults" Grid.Column="2" Style="{StaticResource SecondaryButton}"
                                Width="120" Height="32" Margin="8 0 0 0" Content="Reset Defaults"/>
                        <StackPanel Grid.Column="3" Margin="16 0 0 0" VerticalAlignment="Center">
                            <TextBlock x:Name="TxtRunStatus" FontWeight="SemiBold" FontSize="12"/>
                            <ProgressBar x:Name="RunProgress" Height="6" Minimum="0" Maximum="1" Value="0"
                                         Foreground="#0078D4" Background="#E8E8E8" Margin="0 3 0 0"/>
                        </StackPanel>
                        <TextBlock x:Name="TxtRuntimeSummary" Grid.Column="4" VerticalAlignment="Center"
                                   FontSize="11.5" Foreground="#555" Text="Ready."/>
                    </Grid>
                </Border>

                <!-- Validation -->
                <Border Grid.Row="1" Background="#FFFDF6EC" CornerRadius="4" Padding="8" Margin="0 0 0 4"
                        BorderBrush="#E8D5B0" BorderThickness="1">
                    <TextBlock x:Name="TxtValidation" FontSize="11" Foreground="#6B5B3E"/>
                </Border>

                <!-- Log Output -->
                <TextBox
                    x:Name="TxtLog"
                    Grid.Row="2"
                    AcceptsReturn="True"
                    IsReadOnly="True"
                    VerticalScrollBarVisibility="Auto"
                    HorizontalScrollBarVisibility="Auto"
                    FontFamily="Cascadia Mono,Consolas,Courier New"
                    FontSize="11"
                    Background="#1E1E1E"
                    Foreground="#D4D4D4"
                    BorderBrush="#333"
                    BorderThickness="1"
                    Padding="6"/>
            </Grid>
        </Grid>

        <!-- Status Bar -->
        <Border Grid.Row="2" Background="#F5F5F5" BorderBrush="#D0D0D0" BorderThickness="0 1 0 0" Padding="12 4">
            <TextBlock x:Name="TxtStatusBar" FontSize="11" Foreground="#666" Text="Ready"/>
        </Border>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$MainWindow = [System.Windows.Markup.XamlReader]::Load($reader)

$script:TxtSubscriptionFilter = $MainWindow.FindName("TxtSubscriptionFilter")
$script:TxtTenantFilter = $MainWindow.FindName("TxtTenantFilter")
$script:TxtRegionFilter = $MainWindow.FindName("TxtRegionFilter")
$script:TxtAccountFilter = $MainWindow.FindName("TxtAccountFilter")
$script:BtnDiscover = $MainWindow.FindName("BtnDiscover")
$script:BtnSelectAll = $MainWindow.FindName("BtnSelectAll")
$script:BtnClearAll = $MainWindow.FindName("BtnClearAll")
$script:TxtSummary = $MainWindow.FindName("TxtSummary")
$script:DiscoveryProgress = $MainWindow.FindName("DiscoveryProgress")
$script:TxtStatusHint = $MainWindow.FindName("TxtStatusHint")

$script:RequestGrid = $MainWindow.FindName("RequestGrid")
$script:RunProgress = $MainWindow.FindName("RunProgress")

$script:TxtToken = $MainWindow.FindName("TxtToken")
$script:ChkTryAzCliToken = $MainWindow.FindName("ChkTryAzCliToken")
$script:ChkUseDeviceCode = $MainWindow.FindName("ChkUseDeviceCode")
$script:TxtDelaySeconds = $MainWindow.FindName("TxtDelaySeconds")
$script:TxtRequestsPerMinute = $MainWindow.FindName("TxtRequestsPerMinute")
$script:TxtMaxRetries = $MainWindow.FindName("TxtMaxRetries")
$script:TxtBaseRetrySeconds = $MainWindow.FindName("TxtBaseRetrySeconds")
$script:TxtNewLimit = $MainWindow.FindName("TxtNewLimit")
$script:CmbQuotaType = $MainWindow.FindName("CmbQuotaType")
$script:TxtTitle = $MainWindow.FindName("TxtTitle")
$script:TxtContactFirstName = $MainWindow.FindName("TxtContactFirstName")
$script:TxtContactLastName = $MainWindow.FindName("TxtContactLastName")
$script:TxtContactEmail = $MainWindow.FindName("TxtContactEmail")
$script:TxtTimezone = $MainWindow.FindName("TxtTimezone")
$script:TxtSupportLanguage = $MainWindow.FindName("TxtSupportLanguage")
$script:TxtProxy = $MainWindow.FindName("TxtProxy")
$script:ChkProxyDefault = $MainWindow.FindName("ChkProxyDefault")
$script:ChkRotateFingerprint = $MainWindow.FindName("ChkRotateFingerprint")
$script:ChkDryRun = $MainWindow.FindName("ChkDryRun")
$script:ChkStopOnFirstFailure = $MainWindow.FindName("ChkStopOnFirstFailure")
$script:ChkResumeFromState = $MainWindow.FindName("ChkResumeFromState")
$script:ChkRetryFailedRequests = $MainWindow.FindName("ChkRetryFailedRequests")
$script:TxtRunStatePath = $MainWindow.FindName("TxtRunStatePath")
$script:TxtResultJsonPath = $MainWindow.FindName("TxtResultJsonPath")
$script:TxtResultCsvPath = $MainWindow.FindName("TxtResultCsvPath")
$script:TxtTicketTemplatePath = $MainWindow.FindName("TxtTicketTemplatePath")
$script:TxtRunStatus = $MainWindow.FindName("TxtRunStatus")
$script:BtnRun = $MainWindow.FindName("BtnRun")
$script:BtnCancelRun = $MainWindow.FindName("BtnCancelRun")
$script:BtnReloadDefaults = $MainWindow.FindName("BtnReloadDefaults")

$script:TxtValidation = $MainWindow.FindName("TxtValidation")
$script:TxtRuntimeSummary = $MainWindow.FindName("TxtRuntimeSummary")
$script:TxtLog = $MainWindow.FindName("TxtLog")
$script:TxtStatusBar = $MainWindow.FindName("TxtStatusBar")
$script:RegionListBox = $MainWindow.FindName("RegionListBox")
$script:TxtRegionCount = $MainWindow.FindName("TxtRegionCount")

$script:RequestGrid.ItemsSource = $script:RequestRows

function Add-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)][string]$Level = 'INFO'
    )
    $prefix = "[$(Get-Date -Format 'HH:mm:ss')][$Level] "
    $line = $prefix + $Message
    if ($script:TxtLog) {
        $script:TxtLog.AppendText($line + [Environment]::NewLine)
        $script:TxtLog.ScrollToEnd()
    }
    try {
        if (-not [string]::IsNullOrWhiteSpace($script:LogFilePath)) {
            Add-Content -LiteralPath $script:LogFilePath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch {}
}

function Set-TextValue {
    param(
        [Parameter(Mandatory = $true)]$Target,
        [Parameter(Mandatory = $true)][AllowNull()][AllowEmptyString()][string]$Value
    )
    if ($Target -is [System.Windows.Controls.ComboBox]) {
        for ($i = 0; $i -lt $Target.Items.Count; $i++) {
            if ($Target.Items[$i].Content -eq $Value) { $Target.SelectedIndex = $i; return }
        }
        return
    }
    if ($Target -is [System.Windows.Controls.PasswordBox]) { $Target.Password = $Value; return }
    $Target.Text = $Value
}

function Set-CheckValue {
    param([Parameter(Mandatory = $true)]$Target, [Parameter(Mandatory = $true)][bool]$Value)
    $Target.IsChecked = $Value
}

function Load-DefaultsToControls {
    Set-TextValue -Target $script:TxtDelaySeconds -Value $script:Defaults.DelaySeconds
    Set-TextValue -Target $script:TxtRequestsPerMinute -Value $script:Defaults.RequestsPerMinute
    Set-TextValue -Target $script:TxtMaxRetries -Value $script:Defaults.MaxRetries
    Set-TextValue -Target $script:TxtBaseRetrySeconds -Value $script:Defaults.BaseRetrySeconds
    Set-TextValue -Target $script:TxtNewLimit -Value $script:Defaults.NewLimit
    Set-TextValue -Target $script:CmbQuotaType -Value $script:Defaults.QuotaType
    Set-TextValue -Target $script:TxtTitle -Value $script:Defaults.Title
    Set-TextValue -Target $script:TxtContactFirstName -Value $script:Defaults.ContactFirstName
    Set-TextValue -Target $script:TxtContactLastName -Value $script:Defaults.ContactLastName
    Set-TextValue -Target $script:TxtContactEmail -Value $script:Defaults.ContactEmail
    Set-TextValue -Target $script:TxtTimezone -Value $script:Defaults.PreferredTimeZone
    Set-TextValue -Target $script:TxtSupportLanguage -Value $script:Defaults.PreferredSupportLanguage
    Set-TextValue -Target $script:TxtRunStatePath -Value (Join-Path $script:RootPath "run-state.json")
    Set-TextValue -Target $script:TxtResultJsonPath -Value ""
    Set-TextValue -Target $script:TxtResultCsvPath -Value ""
    Set-TextValue -Target $script:TxtTicketTemplatePath -Value $script:TicketTemplatePath
    Set-CheckValue -Target $script:ChkRotateFingerprint -Value $true
    Set-CheckValue -Target $script:ChkTryAzCliToken -Value $true
    Set-CheckValue -Target $script:ChkResumeFromState -Value $false
    Set-CheckValue -Target $script:ChkRetryFailedRequests -Value $false
}

function Update-SummaryText {
    $total = if ($null -ne $script:RequestRows) { $script:RequestRows.Count } else { 0 }
    $selected = @($script:RequestRows | Where-Object { $_.Selected }).Count
    $script:TxtSummary.Text = "Discovered: $total  |  Selected: $selected"
}

function Get-CurrentGuiProfileSnapshot {
    $selectedRows = @($script:RequestRows | Where-Object { $_.Selected } | Select-Object -ExpandProperty Id)
    $delaySeconds = Convert-ToIntValue -Value $script:TxtDelaySeconds.Text -Default ([int]$script:Defaults.DelaySeconds)
    $requestsPerMinute = Convert-ToIntValue -Value $script:TxtRequestsPerMinute.Text -Default ([int]$script:Defaults.RequestsPerMinute)
    $maxRetries = Convert-ToIntValue -Value $script:TxtMaxRetries.Text -Default ([int]$script:Defaults.MaxRetries)
    $baseRetrySeconds = Convert-ToIntValue -Value $script:TxtBaseRetrySeconds.Text -Default ([int]$script:Defaults.BaseRetrySeconds)
    $newLimit = Convert-ToIntValue -Value $script:TxtNewLimit.Text -Default ([int]$script:Defaults.NewLimit)
    if ($newLimit -lt 1) { $newLimit = [int]$script:Defaults.NewLimit }
    $quotaType = if ([string]::IsNullOrWhiteSpace($script:CmbQuotaType.Text)) { [string]$script:Defaults.QuotaType } else { [string]$script:CmbQuotaType.Text }
    $defaultRunStatePath = Join-Path $script:RootPath "run-state.json"
    $runStatePath = if ([string]::IsNullOrWhiteSpace($script:TxtRunStatePath.Text)) { $defaultRunStatePath } else { [string]$script:TxtRunStatePath.Text }
    $ticketTemplatePath = if ([string]::IsNullOrWhiteSpace($script:TxtTicketTemplatePath.Text)) { $script:TicketTemplatePath } else { [string]$script:TxtTicketTemplatePath.Text }
    $tokenProvided = -not [string]::IsNullOrWhiteSpace($script:TxtToken.Password)
    $tokenSource = if ($tokenProvided) { 'Token' } elseif ([bool]$script:ChkTryAzCliToken.IsChecked) { 'AzureCli' } else { 'None' }

    return [ordered]@{
        profileVersion = 1
        createdAt = (Get-Date).ToString('o')
        tokenSource = $tokenSource
        runSettings = @{
            DelaySeconds = $delaySeconds; RequestsPerMinute = [math]::Max(1, $requestsPerMinute)
            MaxRetries = [math]::Max(0, $maxRetries); BaseRetrySeconds = [math]::Max(1, $baseRetrySeconds)
            RotateFingerprint = [bool]$script:ChkRotateFingerprint.IsChecked; MaxRequests = 0
            TryAzCliToken = [bool]$script:ChkTryAzCliToken.IsChecked
            UseDeviceCodeLogin = [bool]$script:ChkUseDeviceCode.IsChecked
            StopOnFirstFailure = [bool]$script:ChkStopOnFirstFailure.IsChecked
        }
        execution = @{
            DryRun = [bool]$script:ChkDryRun.IsChecked
            RetryFailedRequests = [bool]$script:ChkRetryFailedRequests.IsChecked
            ResumeFromState = [bool]$script:ChkResumeFromState.IsChecked
            CancelSignalPath = ''
        }
        proxy = @{ Url = [string]$script:TxtProxy.Text; UseDefaultCredentials = [bool]$script:ChkProxyDefault.IsChecked }
        resume = @{ RunStatePath = $runStatePath }
        defaults = @{ Region = 'eastus'; QuotaType = $quotaType }
        ticket = @{
            NewLimit = $newLimit; QuotaType = $quotaType
            Title = [string]$script:TxtTitle.Text
            ContactFirstName = [string]$script:TxtContactFirstName.Text
            ContactLastName = [string]$script:TxtContactLastName.Text
            ContactEmail = [string]$script:TxtContactEmail.Text
            PreferredTimeZone = [string]$script:TxtTimezone.Text
            PreferredSupportLanguage = [string]$script:TxtSupportLanguage.Text
            TicketTemplatePath = $ticketTemplatePath
            ResultJsonPath = [string]$script:TxtResultJsonPath.Text
            ResultCsvPath = [string]$script:TxtResultCsvPath.Text
        }
        ui = @{ SelectedRequestIds = @($selectedRows) }
    }
}

function Write-GuiProfileSnapshot {
    param([Parameter(Mandatory = $true)]$Profile)
    $parent = Split-Path -Parent $script:ProfilePath
    if (-not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $Profile | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $script:ProfilePath -Encoding UTF8
}

function Convert-GuiProfileToUnifiedSchema {
    param([Parameter(Mandatory = $true)]$Profile)
    $defaultRunStatePath = Join-Path $script:RootPath "run-state.json"
    return Convert-ProfileToUnifiedSchema `
        -Profile $Profile -TemplateDefaults $script:Defaults `
        -DefaultRunStatePath $defaultRunStatePath -DefaultTicketTemplatePath $script:TicketTemplatePath
}

function Apply-GuiProfileSnapshot {
    param([Parameter(Mandatory = $true)]$Profile)
    $rs = $Profile.runSettings; $ex = $Profile.execution; $px = $Profile.proxy; $re = $Profile.resume; $tk = $Profile.ticket
    Set-TextValue -Target $script:TxtDelaySeconds -Value ([string]$rs.DelaySeconds)
    Set-TextValue -Target $script:TxtRequestsPerMinute -Value ([string]$rs.RequestsPerMinute)
    Set-TextValue -Target $script:TxtMaxRetries -Value ([string]$rs.MaxRetries)
    Set-TextValue -Target $script:TxtBaseRetrySeconds -Value ([string]$rs.BaseRetrySeconds)
    Set-TextValue -Target $script:TxtNewLimit -Value ([string]$tk.NewLimit)
    Set-TextValue -Target $script:CmbQuotaType -Value ([string]$tk.QuotaType)
    Set-TextValue -Target $script:TxtTitle -Value ([string]$tk.Title)
    Set-TextValue -Target $script:TxtContactFirstName -Value ([string]$tk.ContactFirstName)
    Set-TextValue -Target $script:TxtContactLastName -Value ([string]$tk.ContactLastName)
    Set-TextValue -Target $script:TxtContactEmail -Value ([string]$tk.ContactEmail)
    Set-TextValue -Target $script:TxtTimezone -Value ([string]$tk.PreferredTimeZone)
    Set-TextValue -Target $script:TxtSupportLanguage -Value ([string]$tk.PreferredSupportLanguage)
    Set-TextValue -Target $script:TxtRunStatePath -Value ([string]$re.RunStatePath)
    Set-TextValue -Target $script:TxtResultJsonPath -Value ([string]$tk.ResultJsonPath)
    Set-TextValue -Target $script:TxtResultCsvPath -Value ([string]$tk.ResultCsvPath)
    Set-TextValue -Target $script:TxtTicketTemplatePath -Value ([string]$tk.TicketTemplatePath)
    Set-TextValue -Target $script:TxtProxy -Value ([string]$px.Url)
    Set-CheckValue -Target $script:ChkTryAzCliToken -Value ([bool]$rs.TryAzCliToken)
    Set-CheckValue -Target $script:ChkUseDeviceCode -Value ([bool]$rs.UseDeviceCodeLogin)
    Set-CheckValue -Target $script:ChkRotateFingerprint -Value ([bool]$rs.RotateFingerprint)
    Set-CheckValue -Target $script:ChkDryRun -Value ([bool]$ex.DryRun)
    Set-CheckValue -Target $script:ChkStopOnFirstFailure -Value ([bool]$rs.StopOnFirstFailure)
    Set-CheckValue -Target $script:ChkResumeFromState -Value ([bool]$ex.ResumeFromState)
    Set-CheckValue -Target $script:ChkRetryFailedRequests -Value ([bool]$ex.RetryFailedRequests)
    Set-CheckValue -Target $script:ChkProxyDefault -Value ([bool]$px.UseDefaultCredentials)
}

function Load-GuiProfile {
    if (-not (Test-Path -LiteralPath $script:ProfilePath)) { return }
    try {
        $raw = Get-Content -LiteralPath $script:ProfilePath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return }
        $profile = ConvertFrom-Json -InputObject $raw -ErrorAction Stop
        if ($null -eq $profile) { return }
        $converted = Convert-GuiProfileToUnifiedSchema -Profile $profile
        Apply-GuiProfileSnapshot -Profile $converted.Profile
        if ($converted.Migrated) {
            Write-GuiProfileSnapshot -Profile $converted.Profile
            Add-Log "Migrated GUI profile to profileVersion=1 schema." "INFO"
        }
    }
    catch { Add-Log "Unable to load UI profile: $($_.Exception.Message)" "WARN" }
}

function Save-GuiProfile {
    try {
        $profile = Get-CurrentGuiProfileSnapshot
        Write-GuiProfileSnapshot -Profile $profile
    }
    catch { Add-Log "Unable to save UI profile: $($_.Exception.Message)" "WARN" }
}

function Update-Validation {
    $lines = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
        $errors.Add("Azure CLI not found in PATH.")
    } else { $lines.Add("OK: Azure CLI available.") }
    $lines.Add("Engine module loaded.")

    $selected = @()
    if ($script:RequestRows.Count -eq 0) {
        $warnings.Add("Discover requests before running.")
    } else {
        $selected = @($script:RequestRows | Where-Object { $_.Selected })
        if (-not $selected -or $selected.Count -eq 0) { $warnings.Add("Select at least one request row.") }
        else { $lines.Add("Selected: $($selected.Count) row(s).") }
    }

    $tokenProvided = -not [string]::IsNullOrWhiteSpace($script:TxtToken.Password)
    if ($tokenProvided -or $script:ChkTryAzCliToken.IsChecked) { $lines.Add("Token source configured.") }
    else { $errors.Add("Provide token or enable CLI token mode.") }

    $delay = 0; $rpm = 0; $retries = 0; $baseRetry = 0; $newLimit = 0
    if ([int]::TryParse($script:TxtDelaySeconds.Text, [ref]$delay) -and $delay -ge 0) { $lines.Add("DelaySeconds valid.") }
    else { $errors.Add("DelaySeconds must be a non-negative integer.") }
    if ([int]::TryParse($script:TxtRequestsPerMinute.Text, [ref]$rpm) -and $rpm -gt 0) { $lines.Add("Requests/minute valid.") }
    else { $errors.Add("RequestsPerMinute must be greater than zero.") }
    if ([int]::TryParse($script:TxtMaxRetries.Text, [ref]$retries) -and $retries -ge 0) { $lines.Add("MaxRetries valid.") }
    else { $errors.Add("MaxRetries must be 0 or greater.") }
    if ([int]::TryParse($script:TxtBaseRetrySeconds.Text, [ref]$baseRetry) -and $baseRetry -gt 0) { $lines.Add("BaseRetrySeconds valid.") }
    else { $errors.Add("BaseRetrySeconds must be greater than zero.") }
    if ([int]::TryParse($script:TxtNewLimit.Text, [ref]$newLimit) -and $newLimit -gt 0) { $lines.Add("NewLimit valid.") }
    else { $errors.Add("NewLimit must be greater than zero.") }

    if (-not [string]::IsNullOrWhiteSpace($script:TxtContactEmail.Text) -and $script:TxtContactEmail.Text -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $lines.Add("Contact email format valid.")
    } else { $errors.Add("Contact email is required and must be valid.") }

    if (-not [string]::IsNullOrWhiteSpace($script:TxtTicketTemplatePath.Text)) {
        if (-not (Test-Path -LiteralPath $script:TxtTicketTemplatePath.Text)) { $errors.Add("Ticket template path is invalid.") }
        else { try { $null = Load-TicketTemplate -TicketTemplatePath $script:TxtTicketTemplatePath.Text } catch { $errors.Add("Ticket template is invalid: $($_.Exception.Message)") } }
    }

    $quotaType = if ([string]::IsNullOrWhiteSpace($script:CmbQuotaType.Text)) { [string]$script:Defaults.QuotaType } else { [string]$script:CmbQuotaType.Text }
    $requestsForPreflight = New-Object System.Collections.Generic.List[object]
    $requestIndex = 0
    $refreshRows = $false
    foreach ($row in $selected) {
        $requestIndex++
        $subscriptionId = if ([string]::IsNullOrWhiteSpace([string]$row.SubscriptionId)) { "" } else { ([string]$row.SubscriptionId).Trim() }
        $accountName = if ([string]::IsNullOrWhiteSpace([string]$row.AccountName)) { "" } else { ([string]$row.AccountName).Trim() }
        $region = if ([string]::IsNullOrWhiteSpace([string]$row.Region)) { "" } else { ([string]$row.Region).Trim() }
        if ($region -ne [string]$row.Region) { $row.Region = $region; $refreshRows = $true }
        if ([string]::IsNullOrWhiteSpace($subscriptionId)) { $errors.Add("Row #$requestIndex missing Subscription ID."); continue }
        if ([string]::IsNullOrWhiteSpace($accountName)) { $errors.Add("Row #$requestIndex missing Account."); continue }

        $discoveredRegions = Convert-ToStringArray -Value (Get-ObjectMemberValue -Object $row -Name 'DiscoveredRegions')
        $regionCheck = Test-DiscoveryRegionValue -Region $region -DiscoveredRegions $discoveredRegions -DefaultRegion 'eastus'
        if ([string]::IsNullOrWhiteSpace($region) -and -not [string]::IsNullOrWhiteSpace($regionCheck.Region)) {
            $row.Region = $regionCheck.Region; $region = $regionCheck.Region; $refreshRows = $true
        }
        foreach ($re in $regionCheck.Errors) { $errors.Add("Row #${requestIndex}: $re") }
        foreach ($rw in $regionCheck.Warnings) { $warnings.Add("Row #${requestIndex}: $rw") }
        if ([string]::IsNullOrWhiteSpace($region)) { continue }
        $requestsForPreflight.Add([pscustomobject]@{ sub = $subscriptionId; account = $accountName; region = $region; limit = [math]::Max(1, $newLimit); quotaType = $quotaType })
    }
    if ($refreshRows) { $script:RequestGrid.Items.Refresh() }

    if ($selected.Count -gt 0 -and $requestsForPreflight.Count -gt 0) {
        try {
            $preflight = Test-AzureSupportPreFlight `
                -Requests @($requestsForPreflight) -Token $script:TxtToken.Password `
                -TryAzCliToken ([bool]$script:ChkTryAzCliToken.IsChecked) -UseDeviceCodeLogin ([bool]$script:ChkUseDeviceCode.IsChecked) `
                -DelaySeconds $delay -MaxRequests 0 -MaxRetries $retries -BaseRetrySeconds $baseRetry `
                -RotateFingerprint:([bool]$script:ChkRotateFingerprint.IsChecked) -RequestsPerMinute $rpm `
                -ProxyUrl $script:TxtProxy.Text -ProxyUseDefaultCredentials:([bool]$script:ChkProxyDefault.IsChecked) `
                -StopOnFirstFailure ([bool]$script:ChkStopOnFirstFailure.IsChecked)
            foreach ($pe in @($preflight.Errors)) { if (-not [string]::IsNullOrWhiteSpace([string]$pe)) { $errors.Add([string]$pe) } }
            foreach ($pw in @($preflight.Warnings)) { if (-not [string]::IsNullOrWhiteSpace([string]$pw)) { $warnings.Add([string]$pw) } }
        }
        catch { $warnings.Add("Unable to run preflight checks: $($_.Exception.Message)") }
    }

    $script:ValidationState = [ordered]@{ Errors = $errors; Warnings = $warnings; IsValid = $errors.Count -eq 0 }
    foreach ($el in $errors) { $lines.Add("ERROR: $el") }
    foreach ($wl in $warnings) { $lines.Add("WARN: $wl") }

    $selectedForRun = @($script:RequestRows | Where-Object { $_.Selected })
    $canRun = ($script:ValidationState.IsValid -and $selectedForRun.Count -gt 0)
    if ($script:BtnRun) { $script:BtnRun.IsEnabled = [bool]$canRun }
    $script:TxtValidation.Text = ($lines -join [Environment]::NewLine)
    $script:TxtStatusBar.Text = if ($errors.Count -gt 0) { "$($errors.Count) error(s)" } elseif ($warnings.Count -gt 0) { "$($warnings.Count) warning(s)" } else { "Validation passed" }
}

function Update-RequestRowsFromRunState {
    param([Parameter(Mandatory = $true)]$RunState)
    if ($null -eq $RunState -or -not $RunState.requestQueue) { return }
    $pending = 0; $completed = 0; $failed = 0
    for ($i = 0; $i -lt $script:RequestRows.Count; $i++) {
        $row = $script:RequestRows[$i]
        $match = $RunState.requestQueue | Where-Object { $_.subscription -eq $row.SubscriptionId -and $_.account -eq $row.AccountName -and $_.region -eq $row.Region } | Select-Object -First 1
        if ($null -eq $match) { continue }
        $row.Status = [string]$match.status
        if ($match.status -in @('Submitted', 'DryRun', 'Failed')) { $completed++ } else { $pending++ }
        if ($match.status -eq 'Failed') { $failed++ }
    }
    $script:RequestGrid.Items.Refresh()
    $script:RunProgress.Value = [double][Math]::Min($script:RunProgress.Maximum, $completed)
    if ($script:TxtRuntimeSummary) {
        $script:TxtRuntimeSummary.Text = "Done: $completed | Active: $pending | Failed: $failed"
    }
}

function Get-RunStateForUi {
    if ([string]::IsNullOrWhiteSpace($script:RunStatePath) -or -not (Test-Path -LiteralPath $script:RunStatePath)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $script:RunStatePath -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ConvertFrom-Json -InputObject $raw -ErrorAction Stop
    }
    catch { return $null }
}

function New-GridRow {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$SubscriptionId,
        [Parameter(Mandatory = $false)][string]$SubscriptionName,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$TenantId = '',
        [Parameter(Mandatory = $true)][string]$AccountName,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Region = '',
        [Parameter(Mandatory = $false)][string[]]$DiscoveredRegions = @()
    )
    return New-DiscoveryGridRow -Id $Id -SubscriptionId $SubscriptionId -SubscriptionName $SubscriptionName -TenantId $TenantId -AccountName $AccountName -Region $Region -DiscoveredRegions $DiscoveredRegions
}

function Clear-RequestRows {
    $script:RequestRows.Clear()
    Update-SummaryText
    Update-Validation
}

function Start-Discovery {
    if ($script:DiscoveryJob -and $script:DiscoveryJob.State -eq 'Running') { return }
    Clear-RequestRows
    $script:TxtStatusHint.Text = "Running discovery..."
    $script:DiscoveryProgress.IsIndeterminate = $true
    Add-Log "Starting discovery job via ticket engine module."
    $script:DiscoveryJob = Start-Job -Name "AzureTicketDiscovery" -ScriptBlock {
        param([string]$SubscriptionFilter, [string]$TenantFilter, [string]$RegionFilter, [string]$AccountFilter, [string]$EngineModulePath)
        if (-not (Test-Path -LiteralPath $EngineModulePath)) { throw "Ticket engine module not found at '$EngineModulePath'." }
        Import-Module -Name $EngineModulePath -Force -ErrorAction Stop
        Write-Output ([pscustomobject]@{ Type = 'Log'; Level = 'INFO'; Message = 'Discovering Batch accounts via engine module.' })
        $discovered = @(AzureSupport.TicketEngine\Get-AzureSupportDiscoveryRows `
            -SubscriptionFilter $SubscriptionFilter -TenantFilter $TenantFilter `
            -RegionFilter $RegionFilter -AccountFilter $AccountFilter)
        Write-Output ([pscustomobject]@{ Type = 'Summary'; Status = 'done'; Count = $discovered.Count; Message = "Discovery complete. $($discovered.Count) request rows discovered." })
        foreach ($row in $discovered) {
            Write-Output ([pscustomobject]@{
                Type = 'RequestRow'; Id = [string]$row.Id; Selected = [bool]$row.Selected
                SubscriptionId = [string]$row.SubscriptionId; SubscriptionName = [string]$row.SubscriptionName
                TenantId = [string]$row.TenantId; AccountName = [string]$row.AccountName
                Region = [string]$row.Region; DiscoveredRegions = @($row.DiscoveredRegions); Status = [string]$row.Status
            })
        }
    } -ArgumentList $script:TxtSubscriptionFilter.Text, $script:TxtTenantFilter.Text, $script:TxtRegionFilter.Text, $script:TxtAccountFilter.Text, $script:EngineModulePath
}

function Apply-DiscoveryItem {
    param([psobject]$Item)
    if ($Item.Type -ne 'RequestRow') { return }
    $script:RequestRows.Add((New-GridRow -Id $Item.Id -SubscriptionId $Item.SubscriptionId -SubscriptionName $Item.SubscriptionName -TenantId $Item.TenantId -AccountName $Item.AccountName -Region $Item.Region -DiscoveredRegions (Convert-ToStringArray -Value $Item.DiscoveredRegions)))
}

function Update-RequestStatus {
    param([Parameter(Mandatory = $true)][string]$RequestId, [Parameter(Mandatory = $true)][string]$Status, [string]$ErrorMessage = '')
    for ($i = 0; $i -lt $script:RequestRows.Count; $i++) {
        $row = $script:RequestRows[$i]
        if ($row.Id -eq $RequestId) {
            $row.Status = $Status
            $row | Add-Member -NotePropertyName Error -NotePropertyValue $ErrorMessage -Force
            break
        }
    }
    $script:RequestGrid.Items.Refresh()
}

function Get-SelectedRequests { return @($script:RequestRows | Where-Object { $_.Selected }) }

function Start-Run {
    if ($script:RunJob -and $script:RunJob.State -eq 'Running') { return }
    $selected = Get-SelectedRequests
    if (-not $selected -or $selected.Count -eq 0) { [System.Windows.MessageBox]::Show("Select at least one request row before running."); return }
    Update-Validation
    if (-not $script:ValidationState.IsValid) { [System.Windows.MessageBox]::Show("Resolve validation issues before running."); return }

    $templatePath = [string]$script:TxtTicketTemplatePath.Text
    if ([string]::IsNullOrWhiteSpace($templatePath)) { $templatePath = $script:TicketTemplatePath }
    try {
        $script:Template = Load-TicketTemplate -TicketTemplatePath $templatePath
        $script:Defaults = Merge-TemplateDefaults -Template $script:Template
        $script:TicketTemplatePath = $templatePath
    }
    catch {
        $msg = "Unable to load ticket template '$templatePath': $($_.Exception.Message)"
        Add-Log $msg 'ERROR'; [System.Windows.MessageBox]::Show($msg); return
    }

    $runQuotaType = 'LowPriority'
    if ($script:CmbQuotaType.SelectedItem -and -not [string]::IsNullOrWhiteSpace($script:CmbQuotaType.SelectedItem.Content)) {
        $runQuotaType = [string]$script:CmbQuotaType.SelectedItem.Content
    }

    $runSettings = @{
        DelaySeconds = [int]$script:TxtDelaySeconds.Text; RequestsPerMinute = [int]$script:TxtRequestsPerMinute.Text
        MaxRetries = [int]$script:TxtMaxRetries.Text; BaseRetrySeconds = [int]$script:TxtBaseRetrySeconds.Text
        Token = $script:TxtToken.Password; TryAzCliToken = [bool]$script:ChkTryAzCliToken.IsChecked
        UseDeviceCodeLogin = [bool]$script:ChkUseDeviceCode.IsChecked; RotateFingerprint = [bool]$script:ChkRotateFingerprint.IsChecked
        NewLimit = [int]$script:TxtNewLimit.Text; QuotaType = $runQuotaType; Title = $script:TxtTitle.Text
        ContactFirstName = $script:TxtContactFirstName.Text; ContactLastName = $script:TxtContactLastName.Text
        ContactEmail = $script:TxtContactEmail.Text; PreferredContactMethod = $script:Defaults.PreferredContactMethod
        TimeZone = $script:TxtTimezone.Text; SupportLanguage = $script:TxtSupportLanguage.Text; Country = $script:Defaults.Country
        ProxyUrl = $script:TxtProxy.Text; ProxyUseDefaultCredentials = [bool]$script:ChkProxyDefault.IsChecked
        DryRun = [bool]$script:ChkDryRun.IsChecked; StopOnFirstFailure = [bool]$script:ChkStopOnFirstFailure.IsChecked
        ResumeFromState = [bool]$script:ChkResumeFromState.IsChecked; RetryFailedRequests = [bool]$script:ChkRetryFailedRequests.IsChecked
        ResultJsonPath = $script:TxtResultJsonPath.Text; ResultCsvPath = $script:TxtResultCsvPath.Text
        TicketTemplatePath = $script:TxtTicketTemplatePath.Text
        ProblemClassificationId = [string]$script:Template.problemClassificationId
        ServiceId = [string]$script:Template.serviceId; SupportPlanId = [string]$script:Template.supportPlanId
        AdvancedDiagnosticConsent = [string]$script:Template.advancedDiagnosticConsent
        QuotaChangeRequestVersion = [string]$script:Template.quotaChangeRequestVersion
        QuotaChangeRequestSubType = [string]$script:Template.quotaChangeRequestSubType
        QuotaRequestType = [string]$script:Template.quotaRequestType
        DescriptionTemplate = [string]$script:Template.descriptionTemplate
        AcceptLanguage = [string]$script:Template.acceptLanguage; Severity = [string]$script:Template.severity
    }

    $selectedForRun = @()
    foreach ($r in $selected) {
        $selectedForRun += [pscustomobject]@{
            sub = [string]$r.SubscriptionId; account = [string]$r.AccountName
            region = [string]$r.Region; limit = [int]$runSettings.NewLimit; quotaType = [string]$runQuotaType
        }
    }

    $script:RunId = [guid]::NewGuid().ToString()
    $script:CancellationFile = Join-Path $env:TEMP ("AzureSupportTicketRun-" + $script:RunId + ".txt")
    if (Test-Path $script:CancellationFile) { Remove-Item $script:CancellationFile -Force -ErrorAction SilentlyContinue }

    $script:RunStatePath = if ([string]::IsNullOrWhiteSpace($script:TxtRunStatePath.Text)) {
        Join-Path $env:TEMP ("AzureSupportTicketRunState-" + $script:RunId + ".json")
    } else { $script:TxtRunStatePath.Text }
    Set-TextValue -Target $script:TxtRunStatePath -Value $script:RunStatePath

    $script:RunProgress.Maximum = [math]::Max(1, $selected.Count)
    $script:RunProgress.Value = 0
    $script:TxtRunStatus.Text = "Preparing execution..."
    Save-GuiProfile
    Add-Log "Starting ticket execution for $($selected.Count) request(s)."

    $script:BtnRun.IsEnabled = $false; $script:BtnRun.Content = 'Running...'
    $script:BtnCancelRun.IsEnabled = $true

    foreach ($selectedRow in $selected) { Update-RequestStatus -RequestId $selectedRow.Id -Status 'Queued' }
    Update-SummaryText

    $engineModule = $script:EngineModulePath
    $runStatePath = $script:RunStatePath
    $script:RunJob = Start-Job -Name "AzureTicketRun-$($script:RunId)" -ScriptBlock {
        param([object[]]$Requests, [hashtable]$Settings, [string]$CancellationFile, [string]$EngineModulePath, [string]$RunStatePath)
        if (-not (Test-Path -LiteralPath $EngineModulePath)) { throw "Ticket engine module not found at '$EngineModulePath'." }
        Import-Module -Name $EngineModulePath -Force -ErrorAction Stop
        $results = AzureSupport.TicketEngine\Invoke-AzureSupportBatchQuotaRun `
            -Requests $Requests -Token $Settings.Token `
            -DelaySeconds $Settings.DelaySeconds -RequestsPerMinute $Settings.RequestsPerMinute `
            -MaxRetries $Settings.MaxRetries -BaseRetrySeconds $Settings.BaseRetrySeconds -MaxRequests 0 `
            -DryRun:([bool]$Settings.DryRun) -ProxyUrl $Settings.ProxyUrl `
            -ProxyUseDefaultCredentials:([bool]$Settings.ProxyUseDefaultCredentials) `
            -RotateFingerprint:([bool]$Settings.RotateFingerprint) -TryAzCliToken ([bool]$Settings.TryAzCliToken) `
            -UseDeviceCodeLogin ([bool]$Settings.UseDeviceCodeLogin) -StopOnFirstFailure ([bool]$Settings.StopOnFirstFailure) `
            -RunStatePath $RunStatePath -ResumeFromState:([bool]$Settings.ResumeFromState) `
            -RetryFailedRequests:([bool]$Settings.RetryFailedRequests) -CancelSignalPath $CancellationFile `
            -ContactFirstName $Settings.ContactFirstName -ContactLastName $Settings.ContactLastName `
            -PreferredContactMethod $Settings.PreferredContactMethod -PrimaryEmailAddress $Settings.ContactEmail `
            -PreferredTimeZone $Settings.TimeZone -Country $Settings.Country `
            -PreferredSupportLanguage $Settings.SupportLanguage -AcceptLanguage $Settings.AcceptLanguage `
            -ProblemClassificationId $Settings.ProblemClassificationId -ServiceId $Settings.ServiceId `
            -Severity $Settings.Severity -Title $Settings.Title -DescriptionTemplate $Settings.DescriptionTemplate `
            -AdvancedDiagnosticConsent $Settings.AdvancedDiagnosticConsent -Require24X7Response $true `
            -SupportPlanId $Settings.SupportPlanId -QuotaChangeRequestVersion $Settings.QuotaChangeRequestVersion `
            -QuotaChangeRequestSubType $Settings.QuotaChangeRequestSubType -QuotaRequestType $Settings.QuotaRequestType `
            -NewLimit $Settings.NewLimit -TicketTemplatePath $Settings.TicketTemplatePath `
            -ResultJsonPath $Settings.ResultJsonPath -ResultCsvPath $Settings.ResultCsvPath
        if ($results) { $results | ConvertTo-Json -Depth 20 } else { Write-Output '[]' }
    } -ArgumentList $selectedForRun, $runSettings, $script:CancellationFile, $engineModule, $runStatePath
}

function Cancel-Run {
    if ($script:CancellationFile) {
        New-Item -ItemType File -Path $script:CancellationFile -Force -ErrorAction SilentlyContinue | Out-Null
    }
    if ($script:RunJob -and $script:RunJob.State -eq 'Running') {
        Stop-Job -Job $script:RunJob -ErrorAction SilentlyContinue
        Add-Log "Requested cancellation."
    }
}

function Poll-Jobs {
    try {
        if ($script:DiscoveryJob) {
            $outputs = Receive-Job -Job $script:DiscoveryJob
            foreach ($item in $outputs) {
                switch ($item.Type) {
                    'Log' { Add-Log $item.Message $item.Level }
                    'Summary' {
                        Add-Log ($item.Message) 'INFO'
                        $script:TxtStatusHint.Text = "Discovery complete. $($item.Count) request rows."
                    }
                    'RequestRow' { Apply-DiscoveryItem -Item $item }
                }
            }
            if ($script:DiscoveryJob.State -ne 'Running') {
                if ($script:DiscoveryJob.State -eq 'Failed') {
                    $err = $script:DiscoveryJob.ChildJobs[0].JobStateInfo.Reason
                    if ($err) { Add-Log "Discovery failed: $err" "ERROR" }
                    $script:TxtStatusHint.Text = "Discovery failed."
                }
                $script:DiscoveryProgress.IsIndeterminate = $false
                Remove-Job -Job $script:DiscoveryJob
                $script:DiscoveryJob = $null
                Update-SummaryText; Update-Validation
            } else { Update-SummaryText }
        }

        if ($script:RunJob) {
            $runState = Get-RunStateForUi
            if ($runState) {
                Update-RequestRowsFromRunState -RunState $runState
                if (-not [string]::IsNullOrWhiteSpace($runState.status)) { $script:TxtRunStatus.Text = "Status: $($runState.status)" }
                else { $script:TxtRunStatus.Text = "Execution in progress..." }
            }
            if ($script:RunJob.State -ne 'Running') {
                $jobOutput = Receive-Job -Job $script:RunJob -Keep
                if ($jobOutput -and -not [string]::IsNullOrWhiteSpace([string]$jobOutput)) {
                    try {
                        $jsonOutput = [string]$jobOutput[-1]
                        $results = ConvertFrom-Json -InputObject $jsonOutput -ErrorAction Stop
                        if ($results) { $script:TxtRuntimeSummary.Text = "Completed: $($results.Count) result(s)." }
                    } catch {}
                }
                $exitState = $script:RunJob.State
                if ($exitState -eq 'Failed') {
                    $err = $script:RunJob.ChildJobs[0].JobStateInfo.Reason
                    if ($err) { Add-Log "Run failed: $err" "ERROR" }
                    $script:TxtRunStatus.Text = "Run failed. Check logs."
                } else {
                    if ([string]::IsNullOrWhiteSpace($script:TxtRunStatus.Text) -or $script:TxtRunStatus.Text -like 'Executing*') {
                        $script:TxtRunStatus.Text = "Run completed."
                    }
                }
                Remove-Job -Job $script:RunJob
                $script:RunJob = $null
                if ($script:CancellationFile -and (Test-Path $script:CancellationFile)) {
                    Remove-Item $script:CancellationFile -Force -ErrorAction SilentlyContinue
                }
                $script:BtnRun.IsEnabled = $true; $script:BtnRun.Content = 'Run Selected'
                $script:BtnCancelRun.IsEnabled = $false
                $script:TxtStatusHint.Text = "Execution idle."
                Update-Validation
            }
        }

        if ($script:RegionJob -and $script:RegionJob.State -ne 'Running') {
            $regionOutput = Receive-Job -Job $script:RegionJob
            if ($regionOutput) {
                $script:AvailableRegions = @($regionOutput | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $script:RegionListBox.Items.Clear()
                foreach ($rg in $script:AvailableRegions) { $null = $script:RegionListBox.Items.Add($rg) }
                $script:TxtRegionCount.Text = "$($script:AvailableRegions.Count) regions available"
            }
            Remove-Job -Job $script:RegionJob -ErrorAction SilentlyContinue
            $script:RegionJob = $null
        }
    }
    catch {
        try { Add-Log "Poll error: $($_.Exception.Message)" "ERROR" } catch {}
    }
}

$script:PollTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:PollTimer.Interval = [TimeSpan]::FromMilliseconds(350)
$script:PollTimer.Add_Tick({ Poll-Jobs })
$script:PollTimer.Start()

$script:RequestGrid.Add_CellEditEnding({
    param($sender, $e)
    try {
        if ($e.Column.Header -ne 'Region') { return }
        $row = $e.Row.Item
        $editElement = $e.EditingElement
        if ($editElement -is [System.Windows.Controls.TextBox]) {
            $newRegion = $editElement.Text.Trim()
            $discoveredRegions = Convert-ToStringArray -Value (Get-ObjectMemberValue -Object $row -Name 'DiscoveredRegions')
            $regionCheck = Test-DiscoveryRegionValue -Region $newRegion -DiscoveredRegions $discoveredRegions -DefaultRegion 'eastus'
            if (-not [string]::IsNullOrWhiteSpace($regionCheck.Region) -and $regionCheck.Region -ne $newRegion) { $editElement.Text = $regionCheck.Region }
            foreach ($w in $regionCheck.Warnings) { Add-Log "Region edit: $w" 'WARN' }
            foreach ($err in $regionCheck.Errors) { Add-Log "Region edit: $err" 'ERROR' }
        }
        $script:RequestGrid.Dispatcher.BeginInvoke([Action]{ Update-Validation }, [System.Windows.Threading.DispatcherPriority]::Background)
    } catch { Add-Log "Region edit error: $($_.Exception.Message)" "ERROR" }
})

# Numeric-only input filter for TextBoxes
$numericFilter = {
    param($sender, $e)
    if ($e.Text -notmatch '^\d+$') { $e.Handled = $true }
}
foreach ($numBox in @($script:TxtDelaySeconds, $script:TxtRequestsPerMinute, $script:TxtMaxRetries, $script:TxtBaseRetrySeconds, $script:TxtNewLimit)) {
    $numBox.Add_PreviewTextInput($numericFilter)
}

$script:BtnDiscover.Add_Click({ try { Start-Discovery } catch { Add-Log "Discovery error: $($_.Exception.Message)" "ERROR" } })
$script:BtnSelectAll.Add_Click({
    try {
        for ($i = 0; $i -lt $script:RequestRows.Count; $i++) { $script:RequestRows[$i].Selected = $true }
        $script:RequestGrid.Items.Refresh(); Update-SummaryText; Update-Validation
    } catch { Add-Log "Select all error: $($_.Exception.Message)" "ERROR" }
})
$script:BtnClearAll.Add_Click({
    try {
        for ($i = 0; $i -lt $script:RequestRows.Count; $i++) { $script:RequestRows[$i].Selected = $false }
        $script:RequestGrid.Items.Refresh(); Update-SummaryText; Update-Validation
    } catch { Add-Log "Clear error: $($_.Exception.Message)" "ERROR" }
})
$script:BtnRun.Add_Click({ try { Start-Run } catch { Add-Log "Run error: $($_.Exception.Message)" "ERROR" } })
$script:BtnCancelRun.Add_Click({ try { Cancel-Run } catch { Add-Log "Cancel error: $($_.Exception.Message)" "ERROR" } })
$script:BtnReloadDefaults.Add_Click({ try { Load-DefaultsToControls; Update-Validation } catch { Add-Log "Reset error: $($_.Exception.Message)" "ERROR" } })

foreach ($ctrl in @($script:TxtDelaySeconds, $script:TxtRequestsPerMinute, $script:TxtMaxRetries, $script:TxtBaseRetrySeconds, $script:TxtNewLimit, $script:TxtContactEmail, $script:TxtRunStatePath, $script:TxtResultJsonPath, $script:TxtResultCsvPath, $script:TxtTicketTemplatePath)) {
    $ctrl.Add_TextChanged({ try { Update-Validation } catch {} })
}
$script:TxtToken.Add_PasswordChanged({ try { Update-Validation } catch {} })
foreach ($chk in @($script:ChkTryAzCliToken, $script:ChkRotateFingerprint, $script:ChkDryRun)) {
    $chk.Add_Checked({ try { Update-Validation } catch {} })
    $chk.Add_Unchecked({ try { Update-Validation } catch {} })
}

# Region ListBox selection enforcement: max 10
$script:RegionListBox.Add_SelectionChanged({
    param($sender, $e)
    try {
        if ($script:RegionListBox.SelectedItems.Count -gt 10) {
            foreach ($added in $e.AddedItems) { $script:RegionListBox.SelectedItems.Remove($added) }
            Add-Log "Maximum 10 regions can be selected." "WARN"
        }
        $count = $script:RegionListBox.SelectedItems.Count
        $script:TxtRegionCount.Text = "$($script:AvailableRegions.Count) regions available, $count selected"
    } catch {}
})

Load-DefaultsToControls
Load-GuiProfile
Update-Validation
Update-SummaryText
$script:BtnCancelRun.IsEnabled = $false

Add-Log "Loaded ticket engine module: $($script:EngineModulePath)"
Add-Log "Log file: $($script:LogFilePath)"

# Load regions in background
$script:RegionJob = Start-Job -Name "AzureRegionLoad" -ScriptBlock {
    param([string]$EngineModulePath)
    if (Test-Path -LiteralPath $EngineModulePath) {
        Import-Module -Name $EngineModulePath -Force -ErrorAction Stop
        $regions = @(AzureSupport.TicketEngine\Get-AzureRegionList)
        return $regions
    }
    return @('eastus','eastus2','westus','westus2','westus3','centralus','northcentralus','southcentralus',
             'northeurope','westeurope','uksouth','francecentral','germanywestcentral','swedencentral',
             'norwayeast','japaneast','southeastasia','australiaeast','canadacentral','brazilsouth',
             'koreacentral','centralindia')
} -ArgumentList $script:EngineModulePath

[void]$MainWindow.ShowDialog()
