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
    procedure LogError(const Msg: string);  // ������־��¼����
  public
  end;

var
  Form1: TForm1;
  UptimeLabel: TLabel;
  Timer1: TTimer;
  ErrorLabel: TLabel;  // ������ʾ������Ϣ�ı�ǩ

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  // ����Ӳ����Ϣ
  LoadHardwareInfo;

  // ������ʱ��
  Timer1 := TTimer.Create(Self);
  Timer1.Interval := 1000;  // ÿ�����һ��
  Timer1.OnTimer := Timer1Timer;  // �󶨶�ʱ���¼�

  // ����������ʾ������Ϣ�ı�ǩ
  ErrorLabel := TLabel.Create(Self);
  ErrorLabel.Parent := Self;
  ErrorLabel.Left := 10;
  ErrorLabel.Top := 10;
  ErrorLabel.Caption := '';  // Ĭ�ϲ���ʾ������Ϣ
  ErrorLabel.Font.Color := clRed;  // ������Ϣʹ�ú�ɫ��ʾ
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
  PropertyText: string;  // ���ڴ洢ÿ�����Ե��ı�
begin
  InfoList.Clear;
  try
    CoInitialize(nil);  // ��ʼ�� COM ��
    try
      Locator := CreateOleObject('WbemScripting.SWbemLocator');  // ���� WMI Locator ����

      try
        // �������ӵ� WMI ���� (ʹ�� localhost �� 127.0.0.1)
        Services := Locator.ConnectServer('localhost', 'root\CIMV2');
      except
        on E: Exception do
        begin
          LogError('WMI ����ʧ��: ' + E.Message);
          Exit;
        end;
      end;

      // ִ�� WMI ��ѯ
      try
        ObjectSet := Services.ExecQuery('SELECT * FROM ' + WMIClass, 'WQL');
      except
        on E: Exception do
        begin
          LogError('WMI ��ѯʧ��: ' + E.Message);
          Exit;
        end;
      end;

      // ������ѯ���
      Enum := IUnknown(ObjectSet._NewEnum) as IEnumVariant;
      while Enum.Next(1, WMIObject, Value) = 0 do
      begin
        Line := '';
        for i := Low(Properties) to High(Properties) do
        begin
          try
            // �������ֵ�Ƿ�Ϊ�գ��������ֵ���µĴ���
            if not VarIsNull(WMIObject.Properties_.Item(Properties[i]).Value) then
            begin
              propValue := WMIObject.Properties_.Item(Properties[i]).Value;  // ��ȡ����ֵ

              // ����������ƺ����͵�����
              PropertyText := Properties[i] + ': ' + VarToStr(propValue);

              // �� PropertyText ��ֵ����ǩ������� UptimeLabel �� ErrorLabel��
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
              Line := Line + Properties[i] + ': N/A';  // �������ֵΪ�գ���ʾ N/A
          except
            on E: Exception do
              Line := Line + Properties[i] + ': Error reading value';
          end;
        end;
        InfoList.Add(Trim(Line));  // ȷ��ֻ���ַ�������ӵ� InfoList
        WMIObject := Unassigned;
      end;
    finally
      CoUninitialize;  // �ͷ� COM ��Դ
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
  TickCount := GetTickCount div 1000; // ת��Ϊ��
  Result := Format('%d �� %d Сʱ %d ���� %d ��',
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
  Inc(TopOffset, LabelControl.Height + 5); // ����λ��
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

    // ��ȡ��������Ϣ
    GetHardwareInfo('Win32_Processor', ['Name', 'Manufacturer'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('������: ' + InfoList[i], TopOffset);

    // ��ȡ�ڴ���Ϣ
    GetHardwareInfo('Win32_PhysicalMemory', ['Capacity', 'Speed'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('�ڴ�: ' + InfoList[i], TopOffset);

    // ��ȡ�Կ���Ϣ
    GetHardwareInfo('Win32_VideoController', ['Name', 'DriverVersion'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('�Կ�: ' + InfoList[i], TopOffset);

    // ��ȡ������Ϣ
    GetHardwareInfo('Win32_BaseBoard', ['Manufacturer', 'Product'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('����: ' + InfoList[i], TopOffset);

    // ��ȡ��ʾ����Ϣ
    GetHardwareInfo('Win32_DesktopMonitor', ['Name'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('��ʾ��: ' + InfoList[i], TopOffset);

    // ��ȡӲ����Ϣ
    GetHardwareInfo('Win32_DiskDrive', ['Model', 'Size'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('Ӳ��: ' + InfoList[i], TopOffset);

    // ��ȡ������Ϣ
    GetHardwareInfo('Win32_NetworkAdapter', ['Name', 'MACAddress'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('����: ' + InfoList[i], TopOffset);

    // ��ȡ������Ϣ
    GetHardwareInfo('Win32_SoundDevice', ['Name'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('����: ' + InfoList[i], TopOffset);

    // ��ȡ�����Ϣ
    GetHardwareInfo('Win32_Battery', ['Name', 'EstimatedChargeRemaining'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('���: ' + InfoList[i], TopOffset);

    // ��ȡ������Ϣ
    GetHardwareInfo('Win32_CDROMDrive', ['Name'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('����: ' + InfoList[i], TopOffset);

    // ��ȡϵͳ�汾
    GetHardwareInfo('Win32_OperatingSystem', ['Caption', 'Version', 'BuildNumber'], InfoList);
    for i := 0 to InfoList.Count - 1 do
      AddLabel('ϵͳ�汾: ' + InfoList[i], TopOffset);

    // �������ʱ���ǩ
    UptimeLabel := TLabel.Create(Self);
    UptimeLabel.Parent := Self;
    UptimeLabel.Left := 10;
    UptimeLabel.Top := TopOffset;
    UptimeLabel.Caption := '����ʱ��: ' + GetSystemUptime;

  finally
    InfoList.Free;
  end;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  if Assigned(UptimeLabel) then
    UptimeLabel.Caption := '����ʱ��: ' + GetSystemUptime;
end;

procedure TForm1.LogError(const Msg: string);
begin
  ErrorLabel.Caption := '����: ' + Msg;
end;

end.

