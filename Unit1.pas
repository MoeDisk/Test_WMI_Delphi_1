unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ActiveX, ComObj, DateUtils, ExtCtrls;

type
  TForm1 = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    procedure AddLabel(const Text: string; var TopOffset: Integer);
    procedure GetHardwareInfo(WMIClass: string; Properties: array of string; InfoList: TStrings);
    function GetSystemUptime: string;
    procedure LoadHardwareInfo;
    procedure Timer1Timer(Sender: TObject);
    procedure LogError(const Msg: string);  // 错误日志记录方法
  public
  end;

var
  Form1: TForm1;
  UptimeLabel: TLabel;
  Timer1: TTimer;
  ErrorLabel: TLabel;  // 用于显示错误信息的标签

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  // 加载硬件信息
  LoadHardwareInfo;

  // 创建定时器
  Timer1 := TTimer.Create(Self);
  Timer1.Interval := 1000;  // 每秒更新一次
  Timer1.OnTimer := Timer1Timer;  // 绑定定时器事件

  // 创建用于显示错误信息的标签
  ErrorLabel := TLabel.Create(Self);
  ErrorLabel.Parent := Self;
  ErrorLabel.Left := 10;
  ErrorLabel.Top := 10;
  ErrorLabel.Caption := '';  // 默认不显示错误信息
  ErrorLabel.Font.Color := clRed;  // 错误信息使用红色显示
end;

procedure TForm1.GetHardwareInfo(WMIClass: string; Properties: array of string; InfoList: TStrings);
var
  Locator: OLEVariant;
  Services: OLEVariant;
  ObjectSet: OLEVariant;
  WMIObject: OLEVariant;
  Enum: IEnumVariant;
  Value: Cardinal;
  i: Integer;
  Line: string;
  propValue: Variant;
  PropertyText: string;  // 用于存储每个属性的文本
begin
  InfoList.Clear;
  try
    CoInitialize(nil);  // 初始化 COM 库
    try
      Locator := CreateOleObject('WbemScripting.SWbemLocator');  // 创建 WMI Locator 对象

      try
        // 尝试连接到 WMI 服务 (使用 localhost 或 127.0.0.1)
        Services := Locator.ConnectServer('localhost', 'root\CIMV2');
      except
        on E: Exception do
        begin
          LogError('WMI 连接失败: ' + E.Message);
          Exit;
        end;
      end;

      // 执行 WMI 查询
      try
        ObjectSet := Services.ExecQuery('SELECT * FROM ' + WMIClass, 'WQL');
      except
        on E: Exception do
        begin
          LogError('WMI 查询失败: ' + E.Message);
          Exit;
        end;
      end;

      // 遍历查询结果
      Enum := IUnknown(ObjectSet._NewEnum) as IEnumVariant;
      while Enum.Next(1, WMIObject, Value) = 0 do
      begin
        Line := '';
        for i := Low(Properties) to High(Properties) do
        begin
          try
            // 检查属性值是否为空，并避免空值导致的错误
            if not VarIsNull(WMIObject.Properties_.Item(Properties[i]).Value) then
            begin
              propValue := WMIObject.Properties_.Item(Properties[i]).Value;  // 获取属性值

              // 输出属性名称和类型到变量
              PropertyText := Properties[i] + ': ' + VarToStr(propValue);

              // 将 PropertyText 赋值给标签（如果有 UptimeLabel 或 ErrorLabel）
              if Assigned(UptimeLabel) then
                UptimeLabel.Caption := PropertyText;

              case VarType(propValue) of
                varString:
                  Line := Line + PropertyText + ' ';
                varInteger:
                  Line := Line + PropertyText + ' ';
                varDouble:
                  Line := Line + PropertyText + ' ';
                vtInt64:
                  Line := Line + PropertyText + ' ';
                varBoolean:
                  Line := Line + PropertyText + ' ';
                else
                  Line := Line + PropertyText + ' Unknown type';
              end;
            end
            else
              Line := Line + Properties[i] + ': N/A';  // 如果属性值为空，显示 N/A
          except
            on E: Exception do
              Line := Line + Properties[i] + ': Error reading value';
          end;
        end;
        InfoList.Add(Trim(Line));  // 确保只有字符串被添加到 InfoList
        WMIObject := Unassigned;
      end;
    finally
      CoUninitialize;  // 释放 COM 资源
    end;
  except
    on E: Exception do
      LogError('Error: ' + E.Message);
  end;
end;

function TForm1.GetSystemUptime: string;
var
  TickCount: Cardinal;
begin
  TickCount := GetTickCount div 1000; // 转换为秒
  Result := Format('%d 天 %d 小时 %d 分钟 %d 秒',
    [TickCount div 86400, (TickCount div 3600) mod 24, (TickCount div 60) mod 60, TickCount mod 60]);
end;

procedure TForm1.AddLabel(const Text: string; var TopOffset: Integer);
var
  LabelControl: TLabel;
begin
  LabelControl := TLabel.Create(Self);
  LabelControl.Parent := Self;
  LabelControl.Left := 10;
  LabelControl.Top := TopOffset;
  LabelControl.Caption := Text;
  Inc(TopOffset, LabelControl.Height + 5); // 调整位置
end;

procedure TForm1.LoadHardwareInfo;
var
  InfoList: TStringList;
  TopOffset: Integer;
  i: Integer;
begin
  InfoList := TStringList.Create;
  try
    TopOffset := 10;

    // 获取处理器信息
    GetHardwareInfo('Win32_Processor', ['Name', 'Manufacturer'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('处理器: ' + InfoList[i], TopOffset);

    // 获取内存信息
    GetHardwareInfo('Win32_PhysicalMemory', ['Capacity', 'Speed'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('内存: ' + InfoList[i], TopOffset);

    // 获取显卡信息
    GetHardwareInfo('Win32_VideoController', ['Name', 'DriverVersion'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('显卡: ' + InfoList[i], TopOffset);

    // 获取主板信息
    GetHardwareInfo('Win32_BaseBoard', ['Manufacturer', 'Product'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('主板: ' + InfoList[i], TopOffset);

    // 获取显示器信息
    GetHardwareInfo('Win32_DesktopMonitor', ['Name'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('显示器: ' + InfoList[i], TopOffset);

    // 获取硬盘信息
    GetHardwareInfo('Win32_DiskDrive', ['Model', 'Size'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('硬盘: ' + InfoList[i], TopOffset);

    // 获取网卡信息
    GetHardwareInfo('Win32_NetworkAdapter', ['Name', 'MACAddress'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('网卡: ' + InfoList[i], TopOffset);

    // 获取声卡信息
    GetHardwareInfo('Win32_SoundDevice', ['Name'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('声卡: ' + InfoList[i], TopOffset);

    // 获取电池信息
    GetHardwareInfo('Win32_Battery', ['Name', 'EstimatedChargeRemaining'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('电池: ' + InfoList[i], TopOffset);

    // 获取光驱信息
    GetHardwareInfo('Win32_CDROMDrive', ['Name'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('光驱: ' + InfoList[i], TopOffset);

    // 获取系统版本
    GetHardwareInfo('Win32_OperatingSystem', ['Caption', 'Version', 'BuildNumber'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('系统版本: ' + InfoList[i], TopOffset);

    // 添加运行时间标签
    UptimeLabel := TLabel.Create(Self);
    UptimeLabel.Parent := Self;
    UptimeLabel.Left := 10;
    UptimeLabel.Top := TopOffset;
    UptimeLabel.Caption := '运行时间: ' + GetSystemUptime;

  finally
    InfoList.Free;
  end;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  if Assigned(UptimeLabel) then
    UptimeLabel.Caption := '运行时间: ' + GetSystemUptime;
end;

procedure TForm1.LogError(const Msg: string);
begin
  ErrorLabel.Caption := '错误: ' + Msg;
end;

end.

