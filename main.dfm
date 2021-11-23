object UsbNinjaSvc: TUsbNinjaSvc
  OldCreateOrder = False
  DisplayName = 'USB Ninja'
  OnExecute = ServiceExecute
  OnStart = ServiceStart
  OnStop = ServiceStop
  Height = 146
  Width = 487
  object CheckDevices: TTimer
    Enabled = False
    Interval = 1800000
    OnTimer = CheckDevicesTimer
    Left = 40
    Top = 8
  end
  object TimerClearTNetLogLists: TTimer
    Enabled = False
    Interval = 15000
    OnTimer = TimerClearTNetLogListsTimer
    Left = 152
    Top = 8
  end
  object TimerClearEventLogLists: TTimer
    Enabled = False
    Interval = 43200000
    OnTimer = TimerClearEventLogListsTimer
    Left = 280
    Top = 8
  end
end
