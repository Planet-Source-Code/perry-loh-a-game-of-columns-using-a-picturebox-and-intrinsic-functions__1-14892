Attribute VB_Name = "Global"
Public Const BLACK = 0
Public Const WHITE = 7

Public Const GRAY = 0

Public Const BLUE = 1
Public Const GREEN = 2
Public Const CYAN = 3
Public Const RED = 4
Public Const MAGENTA = 5
Public Const YELLOW = 6

Public Const BRIGHT_WHITE = 15
Public Const LIGHT_MAGENTA = 13
Public Const LIGHT_CYAN = 11


' Will use the bottom 4 colors
Public Const LIGHT_BLUE = 9
Public Const LIGHT_GREEN = 10
Public Const LIGHT_RED = 12
Public Const LIGHT_YELLOW = 14

' Sound API constants
Public Const SND_ASYNC = &H1


' Block and playfield constants
Public Const BLOCK_SIZE = 200
Public Const FIELD_WIDTH = 8
Public Const FIELD_HEIGHT = 15

Public Const TIMER_GAME = 400

' Collision constants
Public Const CLD_NONE = 0
Public Const CLD_WALL = 1
Public Const CLD_BLOCK = 2
Public Const CLD_FLOOR = 3

' Orientation of the blocks
Public Const DIR_RIGHT = 0
Public Const DIR_UP = 1
Public Const DIR_LEFT = 2
Public Const DIR_DOWN = 3

' Misc.
Public Const FLASH_TIME = 50

Public Enum enmSoundType
    enmSpin = 0
    enmClear = 1
    enmLevelUp = 2
End Enum

' UDT of each BLOCK, two BLOCKS make a COLUMN
Public Type udtBlock
    X As Integer
    Y As Integer
    Color As Integer
    Falling As Boolean
End Type

' UDT of each field block
Public Type udtFieldBlock
    Block As Boolean
    Color As Integer
End Type

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'Wait a bit
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
