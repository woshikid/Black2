Attribute VB_Name = "Mixer"
Option Explicit
Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Const MAXPNAMELEN = 32
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETLINEINFOF_SOURCE = &H1&
Private Const MIXER_GETLINECONTROLSF_ALL = &H0&
Private Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Private Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Private Const MIXERCONTROL_CONTROLF_UNIFORM = &H1
Private Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2
Private Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Private Type MIXERLINE
    cbStruct As Long               '  size of MIXERLINE structure
    dwDestination As Long          '  zero based destination index
    dwSource As Long               '  zero based source index (if
                                   '  source)
    dwLineID As Long               '  unique line id for mixer device
    fdwLine As Long                '  state/information about line
    dwUser As Long                 '  driver specific information
    dwComponentType As Long        '  component type line connects to
    cChannels As Long              '  number of channels line supports
    cConnections As Long           '  number of connections (possible)
    cControls As Long              '  number of controls at this line
    szShortName As String * 16
    szName As String * 64
    dwType As Long
    dwDeviceID As Long
    wMid  As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
End Type
Private Type MIXERLINECONTROLS
    cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
    dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                           '  MIXER_GETLINECONTROLSF_ONEBYID or
    dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
    cControls As Long      '  count of controls pmxctrl points to
    cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
    pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type
Private Type MIXERCONTROL
    cbStruct As Long           '  size in Byte of MIXERCONTROL
    dwControlID As Long        '  unique control id for mixer device
    dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
    fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
    cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE
                               '  set
    szShortName As String * 8                       ' short name of
                                                    ' control
    szName As String * 32                           ' long name of
                                                    ' control
    lMinimum As Long           '  Minimum value
    lMaximum As Long           '  Maximum value
    reserved(10) As Long       '  reserved structure space
End Type
Private Type MIXERCONTROLDETAILS
    cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
    dwControlID As Long    '  control id to get/set details on
    cChannels As Long      '  number of channels in paDetails array
    item As Long           '  hwndOwner or cMultipleItems
    cbDetails As Long      '  size of _one_ details_XX struct
    paDetails As Long      '  pointer to array of details_XX structs
End Type
Private Type MIXERCONTROLDETAILS_LISTTEXT
    dwParam1 As Long
    dwParam2 As Long
    szName As String * 32
End Type
Private Type MIXERCONTROLDETAILS_BOOLEAN
    fValue As Long
End Type
Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type
Private lineNum As Long
Private hmx As Long
Private mxl As MIXERLINE
Private mxlc As MIXERLINECONTROLS
Private mxc() As MIXERCONTROL
Private cChannels As Long
Private cMultipleItems As Long
Private listText() As MIXERCONTROLDETAILS_LISTTEXT
Private mxcdLT As MIXERCONTROLDETAILS
Private listBool() As MIXERCONTROLDETAILS_BOOLEAN
Private mxcdB As MIXERCONTROLDETAILS
Private mxcdVolume As MIXERCONTROLDETAILS_UNSIGNED
Private mxcdV As MIXERCONTROLDETAILS
Public mixerReady As Boolean
Public mixerList As String

Public Sub getMixerList()
    On Error Resume Next
    Dim i As Long
    Dim name As String
    Dim devNum As Long
    Dim cDevice As Long
    cDevice = mixerGetNumDevs()
    mixerDeinit
    For devNum = 0 To cDevice - 1
        mixerReady = True
        mixerOpen hmx, devNum, 0, 0, 0
        'get line
        mxl.cbStruct = Len(mxl)
        mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN
        mixerGetLineInfo hmx, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE
        If mxl.cControls > 0 Then
            lineNum = mxl.cConnections
            'get linecontrols
            mxlc.cbStruct = Len(mxlc)
            mxlc.dwLineID = mxl.dwLineID
            mxlc.dwControl = 0
            mxlc.cControls = mxl.cControls
            ReDim mxc(mxl.cControls - 1)
            mxlc.cbmxctrl = LenB(mxc(0))
            mxlc.pamxctrl = VarPtr(mxc(0))
            mixerGetLineControls hmx, mxlc, MIXER_GETLINECONTROLSF_ALL
            'find controls list
            For i = 0 To mxl.cControls - 1
                If MIXERCONTROL_CT_CLASS_LIST = (mxc(i).dwControlType And MIXERCONTROL_CT_CLASS_MASK) Then Exit For
            Next i
            If i < mxl.cControls Then
                cChannels = mxl.cChannels
                cMultipleItems = 0
                If MIXERCONTROL_CONTROLF_UNIFORM And mxc(i).fdwControl Then cChannels = 1
                If MIXERCONTROL_CONTROLF_MULTIPLE And mxc(i).fdwControl Then cMultipleItems = mxc(i).cMultipleItems
                'get ready control details of bool
                mxcdB.cbStruct = Len(mxcdB)
                mxcdB.dwControlID = mxc(i).dwControlID
                mxcdB.cChannels = cChannels
                mxcdB.item = cMultipleItems
                ReDim listBool(cChannels * cMultipleItems - 1)
                mxcdB.cbDetails = LenB(listBool(0))
                mxcdB.paDetails = VarPtr(listBool(0))
                'get control details of text
                mxcdLT.cbStruct = Len(mxcdLT)
                mxcdLT.dwControlID = mxc(i).dwControlID
                mxcdLT.cChannels = cChannels
                mxcdLT.item = cMultipleItems
                ReDim listText(cChannels * cMultipleItems - 1)
                mxcdLT.cbDetails = LenB(listText(0))
                mxcdLT.paDetails = VarPtr(listText(0))
                mixerGetControlDetails hmx, mxcdLT, MIXER_GETCONTROLDETAILSF_LISTTEXT
                'deal the list
                For i = 0 To cMultipleItems - 1 Step cChannels
                    name = listText(i).szName
                    name = LeftB(name, InStrB(name, ChrB(0)) - 1)
                    name = StrConv(name, vbUnicode)
                    listText(i).szName = name
                    If name <> "" Then mixerList = mixerList + name + Chr(0)
                Next i
                Exit For
            End If
        End If
        mixerDeinit
    Next devNum
    If mixerList = "" And mixerReady = True Then
        For i = 0 To lineNum - 1
            mxl.dwSource = i
            mixerGetLineInfo hmx, mxl, MIXER_GETLINEINFOF_SOURCE
            name = Left(mxl.szName, InStr(mxl.szName, Chr(0)) - 1)
            listText(i * cChannels).szName = name
            mixerList = mixerList + name + Chr(0)
        Next i
    End If
End Sub

Public Sub setMixer(ByVal line As String)
    On Error Resume Next
    If mixerReady = False Then Exit Sub
    Dim i As Long
    For i = 0 To cChannels * cMultipleItems - 1 Step cChannels
        If line = Trim(listText(i).szName) Then
            listBool(i).fValue = 1
            listBool(i + cChannels - 1).fValue = 1
        Else
            listBool(i).fValue = 0
            listBool(i + cChannels - 1).fValue = 0
        End If
    Next i
    mixerSetControlDetails hmx, mxcdB, MIXER_SETCONTROLDETAILSF_VALUE
    'set volume
    For i = 0 To lineNum - 1
        mxl.dwSource = i
        mixerGetLineInfo hmx, mxl, MIXER_GETLINEINFOF_SOURCE
        If line = Left(mxl.szName, InStr(mxl.szName, Chr(0)) - 1) Then
            mxlc.dwLineID = mxl.dwLineID
            mxlc.dwControl = 0
            mxlc.cControls = mxl.cControls
            ReDim mxc(mxl.cControls - 1)
            mxlc.cbmxctrl = LenB(mxc(0))
            mxlc.pamxctrl = VarPtr(mxc(0))
            mixerGetLineControls hmx, mxlc, MIXER_GETLINECONTROLSF_ALL
            'set details
            mxcdV.cbStruct = Len(mxcdV)
            mxcdV.dwControlID = mxc(0).dwControlID
            mxcdV.cChannels = 1
            mxcdV.item = 0
            mxcdV.cbDetails = LenB(mxcdVolume)
            mxcdV.paDetails = VarPtr(mxcdVolume)
            mxcdVolume.dwValue = mxc(0).lMaximum
            mixerSetControlDetails hmx, mxcdV, MIXER_SETCONTROLDETAILSF_VALUE
            If mxc(0).lMaximum < 2 Then
                mxcdV.dwControlID = mxcdV.dwControlID + 1
                mxcdVolume.dwValue = 65535
                mixerSetControlDetails hmx, mxcdV, MIXER_SETCONTROLDETAILSF_VALUE
            End If
            Exit For
        End If
    Next i
End Sub

Public Sub mixerDeinit()
    On Error Resume Next
    If mixerReady = False Then Exit Sub
    mixerClose hmx
    mixerList = ""
    Erase listText
    Erase listBool
    Erase mxc
    mixerReady = False
End Sub
