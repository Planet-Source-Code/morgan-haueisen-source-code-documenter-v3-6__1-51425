VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' --- API CONSTANTS
'
Private Const C_CHDEVICENAME      As Long = 32&
Private Const C_CHFORMNAME        As Long = 32&
Private Const C_GMEM_FIXED        As Long = &H0
Private Const C_GMEM_MOVEABLE     As Long = &H2
Private Const C_GMEM_ZEROINIT     As Long = &H40
Private Const C_DM_DUPLEX         As Long = &H1000&
Private Const C_DM_ORIENTATION    As Long = &H1&
'
' --- API TYPES DEFINITION
'
Private Type PRINTDLG_TYPE
   lStructSize         As Long
   hWndOwner           As Long
   hDevMode            As Long
   hDevNames           As Long
   hdc                 As Long
   Flags               As Long
   nFromPage           As Integer
   nToPage             As Integer
   nMinPage            As Integer
   nMaxPage            As Integer
   nCopies             As Integer
   hInstance           As Long
   lCustData           As Long
   lpfnPrintHook       As Long
   lpfnSetupHook       As Long
   lpPrintTemplateName As String
   lpSetupTemplateName As String
   hPrintTemplate      As Long
   hSetupTemplate      As Long
End Type

Private Type DEVNAMES_TYPE
   wDriverOffset As Integer
   wDeviceOffset As Integer
   wOutputOffset As Integer
   wDefault      As Integer
   extra         As String * 100
End Type

Private Type DEVMODE_TYPE
   dmDeviceName        As String * C_CHDEVICENAME
   dmSpecVersion       As Integer
   dmDriverVersion     As Integer
   dmSize              As Integer
   dmDriverExtra       As Integer
   dmFields            As Long
   dmOrientation       As Integer
   dmPaperSize         As Integer
   dmPaperLength       As Integer
   dmPaperWidth        As Integer
   dmScale             As Integer
   dmCopies            As Integer
   dmDefaultSource     As Integer
   dmPrintQuality      As Integer
   dmColor             As Integer
   dmDuplex            As Integer
   dmYResolution       As Integer
   dmTTOption          As Integer
   dmCollate           As Integer
   dmFormName          As String * C_CHFORMNAME
   dmUnusedPadding     As Integer
   dmBitsPerPel        As Integer
   dmPelsWidth         As Long
   dmPelsHeight        As Long
   dmDisplayFlags      As Long
   dmDisplayFrequency  As Long
End Type
'
' --- API DECLARATIONS
'
Private Declare Function PrintDialog Lib "COMDLG32.DLL" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'
' --- PUBLIC ENUM
'
Public Enum enuPrinterConstants
   PD_ALLPAGES = &H0
   PD_COLLATE = &H10
   PD_DISABLEPRINTTOFILE = &H80000
   PD_ENABLEPRINTHOOK = &H1000
   PD_ENABLEPRINTTEMPLATE = &H4000
   PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
   PD_ENABLESETUPHOOK = &H2000
   PD_ENABLESETUPTEMPLATE = &H8000
   PD_ENABLESETUPTEMPLATEHANDLE = &H20000
   PD_HIDEPRINTTOFILE = &H100000
   PD_NONETWORKBUTTON = &H200000
   PD_NOPAGENUMS = &H8
   PD_NOSELECTION = &H4
   PD_NOWARNING = &H80
   PD_PAGENUMS = &H2
   PD_PRINTSETUP = &H40
   PD_PRINTTOFILE = &H20
   PD_RETURNDC = &H100
   PD_RETURNDEFAULT = &H400
   PD_RETURNIC = &H200
   PD_SELECTION = &H1
   PD_SHOWHELP = &H800
   PD_USEDEVMODECOPIES = &H40000
   PD_USEDEVMODECOPIESANDCOLLATE = &H40000
   DLG_PRINT = 0
   DLG_PRINTSETUP = &H40
End Enum

Public Enum enuErrorConstants
   cdlCancel = 32755
End Enum
'
' --- PRIVATE VARIABLES
'
Private mlngMinPage          As Long
Private mlngMaxPage          As Long
Private mlngFromPage         As Long
Private mlngToPage           As Long
Private mlngCopies           As Long
Private mlngMaxCopies        As Long
Private mlngOwnerhWnd        As Long
Private mblnOrientationOk    As Boolean
Private mblnShowPrintToFile  As Boolean
Private mblnCancelError      As Boolean
Private mudtFlags            As enuPrinterConstants

'

' -- PUBLIC MEMBERS

Public Property Get AllowOrientation() As Boolean
   AllowOrientation = mblnOrientationOk
End Property

Public Property Let AllowOrientation(ByVal vNewValue As Boolean)
   mblnOrientationOk = vNewValue
End Property

Public Property Get AllowPrintToFile() As Boolean
   AllowOrientation = mblnShowPrintToFile
End Property

Public Property Let AllowPrintToFile(ByVal vNewValue As Boolean)
   mblnShowPrintToFile = vNewValue
End Property

Public Property Get CancelError() As Boolean
   CancelError = mblnCancelError
End Property

Public Property Let CancelError(ByVal vNewValue As Boolean)
   mblnCancelError = CancelError
End Property

Private Sub Class_Initialize()
   
  Dim strPrinterName As String
   
   mlngMinPage = 0
   mlngMaxPage = 0
   mlngFromPage = 0
   mlngToPage = 0
   mlngCopies = 1
   mlngMaxCopies = 1
   mlngOwnerhWnd = 0
   mblnOrientationOk = True
   mblnCancelError = False
   mblnShowPrintToFile = False
   
   On Error Resume Next
   strPrinterName = GetSetting(App.Title, "Options", "Printer", "None")
   On Error GoTo 0
   
End Sub

Private Sub Class_Terminate()
   SaveSetting App.Title, "Options", "Printer", Printer.DeviceName
End Sub

Public Property Get Copies() As Long
   Copies = mlngCopies
End Property

Public Property Let Copies(ByVal vNewValue As Long)
   mlngCopies = IIf(vNewValue > 1, vNewValue, 1)
End Property

Public Property Get C_MaxCopies() As Long
   C_MaxCopies = mlngMaxCopies
End Property

Public Property Let C_MaxCopies(ByVal vNewValue As Long)
   mlngMaxCopies = IIf(vNewValue > 1, vNewValue, 1)
End Property

Public Property Get Flags() As enuPrinterConstants
   Flags = mudtFlags
End Property

Public Property Let Flags(ByVal vNewValue As enuPrinterConstants)
   mudtFlags = Flags
End Property

Property Get FromPage() As Integer
   
   FromPage = mlngFromPage
   
End Property

Property Let FromPage(ByVal vintNewValue As Integer)
   
   vintNewValue = IIf(vintNewValue > 0, vintNewValue, 0)
   mlngFromPage = vintNewValue
   If vintNewValue > mlngToPage Then mlngToPage = vintNewValue
   If vintNewValue > mlngMaxPage Then mlngMaxPage = vintNewValue
   If vintNewValue < mlngMinPage Then mlngMinPage = vintNewValue
   
End Property

Property Get MaxPage() As Integer
   
   MaxPage = mlngMaxPage
   
End Property

Property Let MaxPage(ByVal vintNewValue As Integer)
   
   vintNewValue = IIf(vintNewValue > 0, vintNewValue, 0)
   mlngMaxPage = vintNewValue
   If vintNewValue < mlngToPage Then mlngToPage = vintNewValue
   If vintNewValue < mlngFromPage Then mlngFromPage = vintNewValue
   If vintNewValue < mlngMinPage Then mlngMinPage = vintNewValue
   
End Property

Property Get MinPage() As Integer
   
   MinPage = mlngMinPage
   
End Property

Property Let MinPage(ByVal vintNewValue As Integer)
   
   vintNewValue = IIf(vintNewValue > 0, vintNewValue, 0)
   mlngMinPage = vintNewValue
   If vintNewValue > mlngFromPage Then mlngFromPage = vintNewValue
   If vintNewValue > mlngToPage Then mlngToPage = vintNewValue
   If vintNewValue > mlngMaxPage Then mlngMaxPage = vintNewValue
   
End Property

Public Property Let Owner_hwnd(ByVal vNewValue As Long)
   mlngOwnerhWnd = vNewValue
End Property

Public Function ShowPrinter(ByVal vudtFlags As enuPrinterConstants) As Boolean
   
  Dim udtPrintDlg        As PRINTDLG_TYPE
  Dim udtDevMode         As DEVMODE_TYPE
  Dim udtDevName         As DEVNAMES_TYPE
  Dim lngDevMode         As Long
  Dim lngDevName         As Long
  Dim pntObjPrinter      As Printer
  Dim blnCancel          As Boolean
  Dim strNewPrinterName  As String
   
   If Not mblnShowPrintToFile Then vudtFlags = vudtFlags + PD_HIDEPRINTTOFILE
   
   blnCancel = False
   'Use PrintDialog to get the handle to a memory
   'block with a udtDevMode and udtDevName structures
   With udtPrintDlg
      .lStructSize = Len(udtPrintDlg)
      .hWndOwner = mlngOwnerhWnd
      .Flags = vudtFlags
      .nMinPage = mlngMinPage
      .nFromPage = mlngFromPage
      .nToPage = mlngToPage
      .nMaxPage = mlngMaxPage
      .nCopies = mlngCopies
   End With
   
   'Set the current orientation and duplex setting
   udtDevMode.dmDeviceName = Printer.DeviceName
   udtDevMode.dmSize = Len(udtDevMode)
   udtDevMode.dmFields = C_DM_ORIENTATION Or C_DM_DUPLEX
   udtDevMode.dmOrientation = Printer.Orientation
   
   On Error Resume Next
   udtDevMode.dmDuplex = Printer.Duplex
   On Error GoTo 0
   
   'Allocate memory for the initialization hDevMode structure
   'and copy the settings gathered above into this memory
   udtPrintDlg.hDevMode = GlobalAlloc(C_GMEM_MOVEABLE Or C_GMEM_ZEROINIT, Len(udtDevMode))
   lngDevMode = GlobalLock(udtPrintDlg.hDevMode)
   
   If lngDevMode > 0 Then
      CopyMemory ByVal lngDevMode, udtDevMode, Len(udtDevMode)
      Call GlobalUnlock(lngDevMode)
   End If
   
   'Set the current driver, device, and port name strings
   With udtDevName
      .wDriverOffset = 8
      .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
      .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
      .wDefault = 0
   End With
   
   With Printer
      udtDevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
   End With
   
   'Allocate memory for the initial hDevName structure
   'and copy the settings gathered above into this memory
   udtPrintDlg.hDevNames = GlobalAlloc(C_GMEM_MOVEABLE Or C_GMEM_ZEROINIT, Len(udtDevName))
   lngDevName = GlobalLock(udtPrintDlg.hDevNames)
   
   If lngDevName > 0 Then
      CopyMemory ByVal lngDevName, udtDevName, Len(udtDevName)
      Call GlobalUnlock(lngDevName)
   End If
   
   'Call the print dialog up and let the user make changes
   If PrintDialog(udtPrintDlg) Then
      
      'First get the udtDevName structure.
      lngDevName = GlobalLock(udtPrintDlg.hDevNames)
      CopyMemory udtDevName, ByVal lngDevName, 45
      Call GlobalUnlock(lngDevName)
      
      With udtPrintDlg
         vudtFlags = .Flags
         mlngFromPage = .nFromPage
         mlngToPage = .nToPage
      End With
      GlobalFree udtPrintDlg.hDevNames
      
      'Get the udtDevMode structure and set the printer properties appropriately
      lngDevMode = GlobalLock(udtPrintDlg.hDevMode)
      CopyMemory udtDevMode, ByVal lngDevMode, Len(udtDevMode)
      Call GlobalUnlock(udtPrintDlg.hDevMode)
      GlobalFree udtPrintDlg.hDevMode
      strNewPrinterName = UCase$(left(udtDevMode.dmDeviceName, InStr(udtDevMode.dmDeviceName, Chr$(0)) - 1))
      
      On Error Resume Next
      'Set the printer
      If UCase(Printer.DeviceName) <> strNewPrinterName Then
         For Each pntObjPrinter In Printers
            If UCase$(pntObjPrinter.DeviceName) = strNewPrinterName Then Set Printer = pntObjPrinter
         Next
      End If
      'Set the printer properties modified by the user
      With Printer
         .ColorMode = udtDevMode.dmColor
         .Copies = IIf(udtDevMode.dmCopies > mlngMaxCopies, mlngMaxCopies, udtDevMode.dmCopies)
         .PaperBin = udtDevMode.dmDefaultSource
         .Duplex = udtDevMode.dmDuplex
         If mblnOrientationOk Then .Orientation = udtDevMode.dmOrientation
         .PaperSize = udtDevMode.dmPaperSize
         .PrintQuality = udtDevMode.dmPrintQuality
         .Zoom = udtDevMode.dmScale
      End With
      On Error GoTo 0
      
    Else
      
      GlobalFree udtPrintDlg.hDevMode
      GlobalFree udtPrintDlg.hDevNames
      blnCancel = True
      If mblnCancelError Then Err.Raise cdlCancel, "PrintDialog", "Cancel."
      
   End If
   
   ShowPrinter = blnCancel
   
End Function

Property Get ToPage() As Integer
   
   ToPage = mlngToPage
   
End Property

Property Let ToPage(ByVal vintNewValue As Integer)
   
   vintNewValue = IIf(vintNewValue > 0, vintNewValue, 0)
   mlngToPage = vintNewValue
   If vintNewValue > mlngMaxPage Then mlngMaxPage = vintNewValue
   If vintNewValue < mlngFromPage Then mlngFromPage = vintNewValue
   If vintNewValue < mlngMinPage Then mlngMinPage = vintNewValue
   
End Property

