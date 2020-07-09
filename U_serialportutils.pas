unit U_serialportutils;

interface

uses SysUtils, Classes, StrUtils, U_serialporttypes;

function HexStrToBytes(aStr: string; out aBuffSize: Integer): PByte;
function HexStrToStr(aStr: ansistring): ansistring;
function StrToBytes(const aStr: ansistring; out aBuffSize: Integer): PByte;
function CleanByteStr(const aStr: string): string;
function ByteArrayToStr(aBytes: PByte; aCount: Integer): string;
function ByteArrayToStrMod(aBytes: PByte; aCount: Integer): string;

implementation

function HexStrToBytes(aStr: string; out aBuffSize: Integer): PByte;
begin
  Result := nil;
  aBuffSize := 0;
  aStr := ReplaceStr(aStr, ' ', '');
  aBuffSize := Length(aStr) div 2;
  if (aStr <> '') and (aBuffSize > 0) then
  begin
    GetMem(Result, aBuffSize);
    if (HexToBin(PChar(aStr), Result, aBuffSize) <> aBuffSize) then
    begin
      FreeMem(Result);
    end;
  end;
end;

function HexStrToStr(aStr: ansistring): ansistring;
var
  aBuffSize: Integer;
begin
  Result := '';
  aStr := ReplaceStr(aStr, ' ', '');
  aBuffSize := Length(aStr) div 2;
  if (aStr <> '') and (aBuffSize > 0) then
  begin
    SetLength(Result, aBuffSize);
    if (HexToBin(PAnsiChar(aStr), @Result[1], aBuffSize) <> aBuffSize) then
    begin
      Result := '';
    end;
  end;
end;

function StrToBytes(const aStr: ansistring; out aBuffSize: Integer): PByte;
var
  i: Integer;
begin
  aBuffSize := Length(aStr);
  GetMem(Result, aBuffSize);
  for i := 0 to aBuffSize - 1 do
    Result[i] := byte(aStr[i + 1]);
end;

function CleanByteStr(const aStr: string): string;
var
  i: Integer;
begin
  Result := aStr;
  for i := 1 to Length(aStr) do
    if (CharInSet(aStr[i], [#0 .. #8])) then
      Result[i] := '.';
end;

function ByteArrayToStr(aBytes: PByte; aCount: Integer): string;
var
  i: Integer;
begin
  SetLength(Result, aCount);
  for i := 0 to aCount - 1 do
    Result[i + 1] := char(aBytes[i]);
end;

function ByteArrayToStrMod(aBytes: PByte; aCount: Integer): string;
var
  i: Integer;
begin
  SetLength(Result, aCount);
  for i := 0 to aCount - 1 do
    if (aBytes[i] in [0 .. 8]) then
      Result[i + 1] := '.'
    else
      Result[i + 1] := char(aBytes[i]);
end;

end.
