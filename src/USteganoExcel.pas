unit USteganoExcel;

interface

/// Сокрытие информации в книге Excel (ориентация ячеек)
procedure WriteMSGToWorkbookAngle(
  MSG: ANSIString;
  FileName: string);

/// Чтение информации, сокрытой в книге Excel (оринтация ячеек)
procedure ReadMSGFromWorkbookAngle(
  var MSG: ANSIString;
  FileName: string);

/// Сокрытие информации в книге Excel (скрытый лист)
procedure WriteMSGToWorkbookSecretSheet(
  MSG: ANSIString;
  FileName: string);

/// Чтение информации, сокрытой в книге Excel (скрытый лист)
procedure ReadMSGFromWorkbookSecretSheet(
  var MSG: ANSIString;
  FileName: string);

var
  /// Стартовая колонка, куда пишется скрываемая информация
  BaseCol: ANSIString = 'AQ';
  /// Стартовая строка, куда пишется скрываемая информация
  BaseRow: LongWord = 47867;
  /// Название "секретного" листа
  SecretSheetName: ANSIString = 'SecretSheet';

implementation

uses
  SysUtils, ComObj;

const
  /// Размер скрываемой структуры (пока это байт) в битах
  ValueBitSize = 8;

  /// Получение определённого бита из заданного байта
  /// Нумерация слева направо от 1 до 7
function GetBitByNum(value, num: byte): boolean;
var
  i: byte;
begin
  for i := 1 to ValueBitSize - num do
    value := value shr 1;
  GetBitByNum := odd(value);
end;

/// Установка определённого бита в заданном байте
/// Нумерация слева направо от 1 до 7
function SetBitByNum(value, num, BitValue: byte): byte;
var
  i: byte;
  tmp: byte;
begin
  tmp := 1;
  for i := 1 to ValueBitSize - num do
    tmp := tmp shl 1;
  if BitValue = 1 then
    value := value or tmp;
  if BitValue = 0 then
    value := value and (255 - tmp);
  SetBitByNum := value;
end;

/// Сокрытие информации в книге Excel (ориентация ячеек)
procedure WriteMSGToWorkbookAngle(
  MSG: ANSIString;
  FileName: string);
var
  i, j: word;
  c: byte;
  BitMsg: packed array of boolean;
  Excel: Variant;
begin
  SetLength(
    BitMsg,
    length(MSG) * ValueBitSize + 1);

  /// Преобразование сообщения в набор бит
  for i := 1 to length(MSG) do
  begin
    c := ord(MSG[i]);
    for j := 1 to ValueBitSize do
      BitMsg[(i - 1) * ValueBitSize + j] := GetBitByNum(
        c,
        j);
  end;

  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[FileName];
  /// Записываем набор бит в таблицу
  for i := 0 to length(MSG) * ValueBitSize - 1 do
    if BitMsg[i + 1] then
      Excel.Range[BaseCol + inttostr(BaseRow + i)].Orientation := 1 // +1 градус кодирует единичный бит
    else
      Excel.Range[BaseCol + inttostr(BaseRow + i)].Orientation := -1; // -1 градус кодирует нулевой бит

  begin
    /// Этот код необходим, чтобы при открытии файла курсор устанавливался на ячейку A1,
    /// а не на ячейку [BaseCol BaseRow]
    Excel.Range['a1'].Orientation := 1;
    Excel.Range['a1'].Orientation := 0;
  end;

  Excel.ActiveWorkbook.Save;
  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;

  SetLength(
    BitMsg,
    0);
end;

procedure ReadMSGFromWorkbookAngle(
  var MSG: ANSIString;
  FileName: string);
var
  i, j: word;
  c: byte;
  BitMsg: packed array of boolean;
  Excel: Variant;
  MsgLength: word;
begin
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[FileName];

  /// Определяем длину скрытого сообщения
  MsgLength := 0;
  while (Excel.Range[BaseCol + inttostr(BaseRow + MsgLength)].Orientation = 1) or (Excel.Range[BaseCol + inttostr(BaseRow + MsgLength)].Orientation = -1) do
    MsgLength := MsgLength + 1;
  SetLength(
    BitMsg,
    MsgLength + 1);
  /// Считываем закодированную битую строку
  for i := 1 to MsgLength do
  begin
    if Excel.Range[BaseCol + inttostr(BaseRow + i - 1)].Orientation = 1 then
      BitMsg[i] := true
    else
      BitMsg[i] := false;
  end;
  /// Восстанавливаем из битовой строки сообщение
  for i := 1 to MsgLength div 8 do
  begin
    c := 0;
    for j := 1 to 8 do
      c := SetBitByNum(
        c,
        j,
        byte(BitMsg[(i - 1) * ValueBitSize + j]));
    MSG := MSG + ANSIChar(chr(c));
  end;

  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;

  SetLength(
    BitMsg,
    0);
end;

procedure WriteMSGToWorkbookSecretSheet(
  MSG: ANSIString;
  FileName: string);
var
  i: word;
  tmp: string;
  fl: boolean;
var
  Excel: Variant;
begin
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[FileName];

  fl := false;
  /// Проверяем, существует ли скрытый лист
  for i := 1 to Excel.ActiveWorkbook.Sheets.Count do
  begin
    tmp := Excel.ActiveWorkbook.Sheets.Item[i].Name;
    if Excel.ActiveWorkbook.Sheets.Item[i].Name = SecretSheetName then
    begin
      /// Если лист уже существует, записываем новую информацию
      Excel.ActiveWorkbook.Sheets.Item[i].Range['a1'] := string(MSG);
      fl := true;
      break;
    end;
  end;
  if not fl then
  /// Если скрытого листа нет, создаём его и записываем туда необходимую информацию
  begin
    Excel.ActiveWorkbook.Sheets.Add;
    Excel.Range['a1'] := string(MSG);
    Excel.ActiveWorkbook.ActiveSheet.Name := SecretSheetName;
    Excel.ActiveWorkbook.ActiveSheet.Visible := false;
  end;

  Excel.ActiveWorkbook.Save;
  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;
end;

procedure ReadMSGFromWorkbookSecretSheet(
  var MSG: ANSIString;
  FileName: string);
var
  Excel: Variant;
  i: word;
  tmp: string;
begin
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open[FileName];
  /// Считываем информацию со скрытого листа
  for i := 1 to Excel.ActiveWorkbook.Sheets.Count do
  begin
    tmp := Excel.ActiveWorkbook.Sheets.Item[i].Name;
    if Excel.ActiveWorkbook.Sheets.Item[i].Name = SecretSheetName then
      MSG := ANSIString(Excel.ActiveWorkbook.Sheets.Item[i].Range['a1']);
  end;

  Excel.ActiveWorkbook.Close;
  Excel.Application.Quit;
end;

end.
