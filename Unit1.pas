unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Outlook2000, OleServer, StdCtrls, Gauges, Math, ExtCtrls,
  mapidefs, mapix, mapiutil, dateutils, ComCtrls, mapi;

type
  TForm1 = class(TForm)
    OutlookApplication1: TOutlookApplication;
    Button3: TButton;
    Gauge1: TGauge;
    Mbxfrom: TLabeledEdit;
    Button1: TButton;
    StaticText1: TStaticText;
    StatusBar1: TStatusBar;
    Timer1: TTimer;
    ComboBox1: TComboBox;
    Label1: TLabel;
    ComboBox2: TComboBox;
    Gauge2: TGauge;
    procedure Button3Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ComboBox2Click(Sender: TObject);
    procedure DoAgenda(Mailbox: String);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}
function pByteToStr(vec: PByte; siz: Integer): string;
var
  text: PChar;
  tsiz: Integer;
begin
  tsiz := siz * 2 + 4; // ta i lite i överkant
  text := StrAlloc(tsiz);
  try
    FillChar(text^, tsiz, 0);
    BinToHex(PChar(vec), text, siz);
    Assert(StrLen(text) = Cardinal(2 * siz));
    Result := text
  finally
    StrDispose(text)
  end
end;


procedure BytesToHexStr(var hHexStr: String; pbyteArray: PByte; InputLength:
Integer);
Const
    HexChars : Array[0..15] of Char = '0123456789ABCDEF';
var
    i, j: WORD;
begin
 SetLength(hHexStr, (InputLength * 2));
 FillChar(hHexStr, sizeof(hHexStr), #0);
 j := 1;
 for i := 1 to InputLength  do begin
  hHexStr[j] := Char(HexChars[pbyteArray^ shr  4]); inc(j);
  hHexStr[j] := Char(HexChars[pbyteArray^ and 15]); inc(j);
  inc(pbyteArray);
 end;
end;

procedure HexBytesToChar(var Response: String; hexbytes: PChar; InputLength:
WORD);
var
 i: WORD;
 c: byte;
begin
 SetLength(Response, InputLength);
 FillChar(Response, SizeOf(Response), #0);
 for i := 0 to (InputLength - 1) do begin
   c := BYTE(hexbytes[i]) And BYTE($f);
   if c > 9 then
     Inc(c, $37)
   else
     Inc(c, $30);
   Response[i + 1] := char(c);
 end;{for}
end;

procedure HexStrToBytes(hHexStr: String; pbyteArray: Pointer);
{pbyteArray must point to enough memory to hold the output}
var
 i, j: WORD;
 tempPtr: PChar;
 twoDigits : String[2];
begin
 tempPtr := pbyteArray;
 j := 1;
 for i := 1 to (Length(hHexStr) DIV 2) do begin
   twoDigits := Copy(hHexStr, j, 2); Inc(j, 2);
   PByte(tempPtr)^ := StrToInt('$' + twoDigits); Inc(tempPtr);
 end;{for}
end;

procedure TForm1.DoAgenda(Mailbox: String);
const
      // GUID of the TimeZone property of recurring appointments
      // as found using OutlookSpy
      PropGUID: TGUID = '{00062002-0000-0000-C000-000000000046}';

      // Correct Timezonestring for DTS, TZ Amsterdam, Berlin, Rome GMT+1
      // Find yours in registry, HKLM\System\CurrentControlSet\Control\TimeZoneInformation
      //                BIAS            Daylight BIAS   StandardStart                   DaylightStart
      //               |--------------||--------------||------------------------------||------------------------------|
      TZStr: String = 'C4FFFFFF00000000C4FFFFFF000000000A00000005000300000000000000000000000300000005000200000000000000';

var i: Integer;
  Recip: Recipient;
  oFolder: MAPIFolder;
  oNs: NameSpace;
  ItemList: Items;
  Appt: AppointmentItem;
  MI: MailItem;
  Msg: IMessage;
  Prop:array[0..0] of TMAPINAMEID;
  pTag:PSPropTagArray;
  Names:array[0..0] of PMAPINAMEID;
  pProp: PSPropValue;
  PR_YOUR_PROP: ULONG;
  rYear, rMonth, rDay, rHour, rMin, rSec, rMSec :Word;
  rK, rIn, rUit: Integer;
  rTIn, rTUit: TDateTime;
  CheckPoint: TDateTime;
  CheckFound: Boolean;
  Time1, Time2: TDateTime;
  Shift: Integer;

begin
  Shift := StrToInt(((ComboBox1.Items.Strings[ComboBox1.ItemIndex])));
  SetRoundMode(rmUp);
  Time1 := Now;

  // Set Checkpoint date to 1-1-2010 01:00
  CheckPoint := EncodeDateTime(2010,1,1,1,0,0,0);

  StatusBar1.SimpleText := 'Outlook wordt geopend';
  OutlookApplication1.Connect;

  StatusBar1.SimpleText := 'Verbinding maken met Outlook';
  oNs := OutlookApplication1.GetNamespace('MAPI');

  StatusBar1.SimpleText := 'Openen Postbus - ' + Mailbox;
  Recip := oNs.CreateRecipient(Mailbox);

  if Recip.Resolve then begin
    MI := OutlookApplication1.CreateItem(olMailItem) as MailItem;
    MI.Recipients.Add('timezone');
    MI.Subject := Recip.Name + ' (start)';
    MI.Body := 'started at: ' + DateTimeToStr(Time1);
    MI.Send;
    MI := nil;

    StatusBar1.SimpleText := 'Openen Agenda';
    try
     oFolder := oNs.GetSharedDefaultFolder(Recip, olFolderCalendar);
    ItemList := oFolder.Items;
    ItemList.IncludeRecurrences := False;
    ItemList.Sort('Start', True);

    StatusBar1.SimpleText := 'Controleren Agenda';

    CheckFound := False;
    Gauge1.MaxValue := ItemList.Count;
    ItemList.Sort('Start', True);

    for i := 1 to ItemList.Count do begin
      Gauge1.Progress := i;
      Application.ProcessMessages;

      if SysUtils.Supports(ItemList.Item(i), AppointmentItem, Appt) then begin
        Appt := ItemList.Item(i) as AppointmentItem;

        // See if checkpoint exists
        if (Appt.Start = Checkpoint) and (Appt.Subject = 'erasmusbrug') then begin
          // ShowMessage('Agenda is reeds aangepast.');
          Time2 := Now;
          MI := OutlookApplication1.CreateItem(olMailItem) as MailItem;
          MI.Recipients.Add('timezone');
          MI.Subject := Recip.Name + ' (finished)';
          MI.Body := 'checkpoint found, nothing was done'#13 + 'processing finished at: ' + DateTimeToStr(Time2);
          MI.Body := MI.Body + #13 + 'requested shift: ' + IntToStr(Shift);
          MI.Body := MI.Body + #13 + 'total progress: ' + IntToStr(Gauge2.PercentDone) + '%';
          MI.Send;
          MI := nil;

          CheckFound := True;
        end;

      end;

      if CheckFound then Break;

    end;

    if not CheckFound then begin
      ItemList.Sort('Start', False);
      // Set MaxValue of gauge to nr of available items
      Gauge1.MaxValue := ItemList.Count;

      StatusBar1.SimpleText := 'Verwerken Agenda';

      for i := 1 to ItemList.Count do begin
        Gauge1.Progress := i;
        Gauge1.Hint := 'Verwerken item ' + IntToStr(i) + ' van ' + IntToStr(ItemList.Count);
        Application.ProcessMessages;

        if SysUtils.Supports(ItemList.Item(i), AppointmentItem, Appt) then begin
          Appt := ItemList.Item(i) as AppointmentItem;

          // Appointment is recurring?
          if Appt.IsRecurring then begin
            Msg:=Appt.MAPIOBJECT as IMessage;

            with Prop[0] do begin
              lpguid:=@propGUID;
              ulKind:=MNID_ID;
              Kind.lID:=$8233; // Found using OutlookSpy
            end;

            Names[0]:=@Prop[0];

            if (S_OK = Msg.GetIDsFromNames(1, Names[0], MAPI_CREATE, PSPropTagArray(pTag))) then begin
              PR_YOUR_PROP:=pTag^.aulPropTag[0] or PT_BINARY ; //is it PT_BINARY?

              if S_OK = HrGetOneProp(Msg, PR_YOUR_PROP, pProp) then begin
                // pProp.Value.bin.lpb now points to the data. pProp.Value.bin.cb is its length
                HexStrToBytes(TZStr, pProp.Value.bin.lpb);

                // Modify the property
                if S_OK = HrSetOneProp(Msg, pProp) then begin
                  Msg.SaveChanges(KEEP_OPEN_READWRITE);
                end;

                MAPIFreeBuffer(pProp);
                MAPIFreeBuffer(pTag);
                Application.ProcessMessages;

              end;

            end;

            Msg:=nil;

          end
          else begin
            // Get the year the appointment starttime is in
            DecodeDate(Appt.Start, rYear, rMonth, rDay);
            DecodeTime(Appt.Start, rHour, rMin, rSec, rMSec);

           (* The following "integer arithmetic" formulae
              will get you the dates when EU Summer Time
              starts and ends, given the year of interest.
              They're guaranteed to work through the year
              2099, or until the EU nations change the rules,
              whichever comes first.

              k = (year * 5) / 4 ["k" and "year" are integers]

              March "Summer Time starts" date = (32762 - k) mod 7  +  25

              October "Summer Time ends" date = (32765 - k) mod 7  +  25

              The value "k" is evaluated by multiplying the
              integer year by 5, then dividing by 4;
              any fraction remaining is discarded -- that is,
              the result is always "floored."

              Example: Compute the Summer Time transition dates for 1998.

              k = (1998 * 5) / 4 = 2497

              March date   = (32762 - 2497) mod 7  +  25 = 29

              October date = (32765 - 2497) mod 7  +  25 = 25

              Hence, the dates of interest are 29 March and 25 October 1998.
           *)

            rK := Round(Int((rYear*5) / 4));
            rIn := (32762 - rK) mod 7 + 25;
            rUit := (32765 - rK) mod 7 + 25;
            rTin := EncodeDateTime(rYear, 3, rIn,2,1,0,0);
            rTuit := EncodeDateTime(rYear, 10, rUit,3,1,0,0);

            // Is Appointment in DTS?
            if (CompareDate(Appt.Start, rTIn) > -1) and (CompareDate(Appt.Start, rTuit) = -1) then begin
                Appt.Start := (Appt.Start + (shift/24));
                Appt.Save;
            end;

          end;

        end;

      end;

      // Ready, create a dummy appointment in the future so we have a checkpoint
      Appt := nil;
      Appt := ItemList.Add(olAppointmentItem) as AppointmentItem;
      Appt.Subject := 'erasmusbrug';
      Appt.Body := 'NIET VERWIJDEREN! deze afspraak is gebruikt voor de omzetting naar zomertijd';
      Appt.Start := CheckPoint;
      Appt.End_ := CheckPoint + (1/24);
      Appt.ReminderSet := False;
      Appt.Save;

      Time2 := Now;
      MI := OutlookApplication1.CreateItem(olMailItem) as MailItem;
      MI.Recipients.Add('timezone');
      MI.Subject := Recip.Name + ' (finished)';
      MI.Body := IntToStr(ItemList.Count) + ' appointments in calendar.'#13 + 'processing finished at: ' + DateTimeToStr(Time2);
      MI.Body := MI.Body + #13 + 'requested shift: ' + IntToStr(Shift);
      MI.Body := MI.Body + #13 + 'total progress: ' + IntToStr(Gauge2.PercentDone) + '%';
      MI.Send;
      MI := nil;

    end;

  Gauge1.Progress := Gauge1.MaxValue;
   except
    MI := OutlookApplication1.CreateItem(olMailItem) as MailItem;
    MI.Recipients.Add('timezone');
    MI.Subject := 'Error: ' + Recip.Name + ' unable to open calendar';
    MI.Body := 'started at: ' + DateTimeToStr(Time1);
    MI.Send;
    MI := nil;
   end;

  end
  else begin
    MI := OutlookApplication1.CreateItem(olMailItem) as MailItem;
    MI.Recipients.Add('timezone');
    MI.Subject := 'Error: ' + Recip.Name + ' did not resolve!';
    MI.Body := 'started at: ' + DateTimeToStr(Time1);
    MI.Send;
    MI := nil;
  end;

  // Garbage Collection
  Appt := nil;
  ItemList := nil;
  Recip := nil;
  oNs.Logoff;
  oNS := nil;
  OutlookApplication1.Disconnect;
  Gauge1.Hint := '';
  StatusBar1.SimpleText := 'Gereed';

end;

Procedure TForm1.Button3Click(Sender: TObject);
begin
  DoAgenda(MbxFrom.Text);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
   OutlookApplication1 := nil;
   MAPIUninitialize;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
//  SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0,
//               SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE);

  Left:=(Screen.Width-Width)  div 2;
  Top:=(Screen.Height-Height) div 2;
  MAPIInitialize(nil);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  if MessageDlg('Are you sure you wish to exit?'#10#13'Note that all running tasks will be terminated.', mtConfirmation, [mbYes,mbNo],0) = mrYes then begin
    Application.Terminate;
  end
  else begin
    {do nothing}
  end;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
var teller: integer;
    lines: LongInt;
    mbxfile: TextFile;
    domailbox: string;
begin
  Timer1.Enabled := False;

  if ParamStr(1) <> '' then begin
    StaticText1.Caption := 'Fix Calendar is running in unattended mode.';
    Button3.Visible := False;
    mbxfrom.Enabled := False;

    if ParamStr(2) = '' then begin
      ShowMessage('No shift specified, therefore assuming -1');
    end
    else begin
      AssignFile(mbxfile, ParamStr(1));
      Reset(mbxfile);

      lines := 0;
      while not eof(mbxfile) do begin
        ReadLn(mbxfile, domailbox);
        Inc(lines);
      end;
      Gauge2.MaxValue := Lines;

      Reset(mbxfile);
      lines := 0;
      while not eof(mbxfile) do begin
        ReadLn(mbxfile, domailbox);
        inc(lines);
        Gauge2.Progress := Lines;
//        ShowMessage(DoMailBox);
        DoAgenda(DoMailBox);
      end;

    end;

    if not FileExists(ParamStr(1)) then begin
      ShowMessage('File not found: ' + ParamStr(1));
      Application.Terminate;
    end
    else begin
      for teller := 1 to ComboBox1.Items.Count do begin
        if ComboBox1.Items.Strings[teller] = ParamStr(2) then begin
          ComboBox1.ItemIndex := teller;
        end;
      end;
    end;
  end
  else begin
    StaticText1.Caption := 'Fix Calendar is running in attended mode.'#10#13'Please type mailbox name and press Start.';
  end;

end;

procedure TForm1.ComboBox2Click(Sender: TObject);
var
  lpRecip: TMapiRecipDesc;
  intRecips: ULONG;
  lpRecips: PMapiRecipDesc;
  i: Integer;
begin
  ComboBox2.Text := '';
  if (MAPIAddress(0, 0, 'Select the recipients to be processed', 1, '', 0, lpRecip, 0, 0, @intRecips, lpRecips) = SUCCESS_SUCCESS) then
  begin
    for i := 0 to intRecips-1 do
      ComboBox2.Items.Add(PMapiRecipDesc(PChar(lpRecips) + i*SizeOf(TMapiRecipDesc))^.lpszName);
    MAPIFreeBuffer(lpRecips)
  end;
  ComboBox2.ItemIndex := 0;
  Application.ProcessMessages;
end;

end.
