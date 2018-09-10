# FixCal
Very old code that fixed up the Outlook timezone identifier for each mailbox on an Exchange server

Probably code to be ashamed of in 2018 but it was written to fixup a problem we had after migrating from Schedule+ to Exchange.
Basically all recurring appointments had a wrong timezone which I reversed to be:
      // GUID of the TimeZone property of recurring appointments
      // as found using OutlookSpy
      PropGUID: TGUID = '{00062002-0000-0000-C000-000000000046}';

      // Correct Timezonestring for DTS, TZ Amsterdam, Berlin, Rome GMT+1
      // Find yours in registry, HKLM\System\CurrentControlSet\Control\TimeZoneInformation
      //                BIAS            Daylight BIAS   StandardStart                   DaylightStart
      //               |--------------||--------------||------------------------------||------------------------------|
      TZStr: String = 'C4FFFFFF00000000C4FFFFFF000000000A00000005000300000000000000000000000300000005000200000000000000';
