Attribute VB_Name = "VideoPortal"
'* Copyright (c) 1996-2000 Logitech, Inc.  All Rights Reserved
'* User Interface Element, codes used with EnableUIElement method
Public Const UIELEMENT_640x480 = 0
Public Const UIELEMENT_320x240 = 1
Public Const UIELEMENT_PCSMART = 2
Public Const UIELEMENT_STATUSBAR = 3
Public Const UIELEMENT_UI = 4
Public Const UIELEMENT_CAMERA = 5
Public Const UIELEMENT_160x120 = 6

'* Camera status codes, returned by CameraState property
Public Const CAMERA_OK = 0
Public Const CAMERA_UNPLUGGED = 1
Public Const CAMERA_INUSE = 2
Public Const CAMERA_ERROR = 3
Public Const CAMERA_SUSPENDED = 4
Public Const CAMERA_DUAL_DETACHED = 5
Public Const CAMERA_UNKNOWNSTATUS = 10

'* Movie Recording Modes, used with MovieRecordMode property
Public Const SEQUENCECAPTURE_FPS_USERSPECIFIED = 1
Public Const SEQUENCECAPTURE_FPS_FASTASPOSSIBLE = 2
Public Const STEPCAPTURE_MANUALTRIGGERED = 3

'* Movie Creation Flags, used with MovieCreateFlags property
Public Const MOVIECREATEFLAGS_CREATENEW = 1
Public Const MOVIECREATEFLAGS_APPEND = 2

'* Notification Codes
Public Const NOTIFICATIONMSG_MOTION = 1
Public Const NOTIFICATIONMSG_MOVIERECORDERROR = 2
Public Const NOTIFICATIONMSG_CAMERADETACHED = 3
Public Const NOTIFICATIONMSG_CAMERAREATTACHED = 4
Public Const NOTIFICATIONMSG_IMAGESIZECHANGE = 5
Public Const NOTIFICATIONMSG_CAMERAPRECHANGE = 6
Public Const NOTIFICATIONMSG_CAMERACHANGEFAILED = 7
Public Const NOTIFICATIONMSG_POSTCAMERACHANGED = 8
Public Const NOTIFICATIONMSG_CAMERBUTTONCLICKED = 9
Public Const NOTIFICATIONMSG_VIDEOHOOK = 10
Public Const NOTIFICATIONMSG_SETTINGDLGCLOSED = 11
Public Const NOTIFICATIONMSG_QUERYPRECAMERAMODIFICATION = 12
Public Const NOTIFICATIONMSG_MOVIESIZE = 13

'* Error codes used by NOTIFICATIONMSG_MOVIERECORDERROR notification:
Public Const WRITEFAILURE_RECORDINGSTOPPED = 0
Public Const WRITEFAILURE_RECORDINGSTOPPED_FILECORRUPTANDDELETED = 1
Public Const WRITEFAILURE_CAMERA_UNPLUGGED = 2
Public Const WRITEFAILURE_CAMERA_SUSPENDED = 3

'* Camera type codes, returned by GetCameraType method
Public Const CAMERA_UNKNOWN = 0
Public Const CAMERA_QUICKCAM_VC = 1
Public Const CAMERA_QUICKCAM_QUICKCLIP = 2
Public Const CAMERA_QUICKCAM_PRO = 3
Public Const CAMERA_QUICKCAM_HOME = 4
Public Const CAMERA_QUICKCAM_PRO_B = 5
Public Const CAMERA_QUICKCAM_TEKCOM = 6
Public Const CAMERA_QUICKCAM_EXPRESS = 7
Public Const CAMERA_QUICKCAM_FROG = 8    '* MIGHT CHANGE NAME BUT ENUM STAYS THE SAME
Public Const CAMERA_QUICKCAM_EMERALD = 9    '* MIGHT CHANGE NAME BUT ENUM STAYS THE SAME

'* Camera-specific property codes used by Set/GetCameraPropertyLong
Public Const PROPERTY_ORIENTATION = 0
Public Const PROPERTY_BRIGHTNESSMODE = 1
Public Const PROPERTY_BRIGHTNESS = 2
Public Const PROPERTY_CONTRAST = 3
Public Const PROPERTY_COLORMODE = 4
Public Const PROPERTY_REDGAIN = 5
Public Const PROPERTY_BLUEGAIN = 6
Public Const PROPERTY_SATURATION = 7
Public Const PROPERTY_EXPOSURE = 8
Public Const PROPERTY_RESET = 9
Public Const PROPERTY_COMPRESSION = 10
Public Const PROPERTY_ANTIBLOOM = 11
Public Const PROPERTY_LOWLIGHTFILTER = 12
Public Const PROPERTY_IMAGEFIELD = 13
Public Const PROPERTY_HUE = 14
Public Const PROPERTY_PORT_TYPE = 15
Public Const PROPERTY_PICTSMART_MODE = 16
Public Const PROPERTY_PICTSMART_LIGHT = 17
Public Const PROPERTY_PICTSMART_LENS = 18
Public Const PROPERTY_MOTION_DETECTION_MODE = 19
Public Const PROPERTY_MOTION_SENSITIVITY = 20
Public Const PROPERTY_WHITELEVEL = 21
Public Const PROPERTY_AUTO_WHITELEVEL = 22
Public Const PROPERTY_ANALOGGAIN = 23
Public Const PROPERTY_AUTO_ANALOGGAIN = 24
Public Const PROPERTY_LOWLIGHTBOOST = 25
Public Const PROPERTY_COLORBOOST = 26
Public Const PROPERTY_ANTIFLICKER = 27
Public Const PROPERTY_OPTIMIZATION_SPEED_QUALITY = 28
Public Const PROPERTY_STREAM_HOOK = 29
Public Const PROPERTY_LED = 30


Public Const ADJUSTMENT_MANUAL = 0
Public Const ADJUSTMENT_AUTOMATIC = 1

Public Const ORIENTATIONMODE_NORMAL = 0
Public Const ORIENTATIONMODE_MIRRORED = 1
Public Const ORIENTATIONMODE_FLIPPED = 2
Public Const ORIENTATIONMODE_FLIPPED_AND_MIRRORED = 3

Public Const COMPRESSION_Q0 = 0
Public Const COMPRESSION_Q1 = 1
Public Const COMPRESSION_Q2 = 2

Public Const ANTIFLICKER_OFF = 0
Public Const ANTIFLICKER_50Hz = 1
Public Const ANTIFLICKER_60Hz = 2

Public Const OPTIMIZE_QUALITY = 0
Public Const OPTIMIZE_SPEED = 1

Public Const LED_OFF = 0
Public Const LED_ON = 1
Public Const LED_AUTO = 2
Public Const LED_MAX = 3

Public Const PICTSMART_LIGHTCORRECTION_NONE = 0
Public Const PICTSMART_LIGHTCORRECTION_COOLFLORESCENT = 1
Public Const PICTSMART_LIGHTCORRECTION_WARMFLORESCENT = 2
Public Const PICTSMART_LIGHTCORRECTION_OUTSIDE = 3
Public Const PICTSMART_LIGHTCORRECTION_TUNGSTEN = 4

Public Const PICTSMART_LENSCORRECTION_NORMAL = 0
Public Const PICTSMART_LENSCORRECTION_WIDEANGLE = 1
Public Const PICTSMART_LENSCORRECTION_TELEPHOTO = 2

Public Const CAMERADLG_GENERAL = 0
Public Const CAMERADLG_ADVANCED = 1





