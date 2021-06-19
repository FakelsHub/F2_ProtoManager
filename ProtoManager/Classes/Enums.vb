﻿Class Enums

    Enum Flags As Integer
'       Mouse_3d     = &H00000001
'       WalkThru     = &H00000004
        Flat         = &H00000008
        NoBlock      = &H00000010
        Lighting     = &H00000020
'       Temp         = &H00000400
        MultiHex     = &H00000800
        NoHighlight  = &H00001000
'       Used         = &H00002000
        TransRed     = &H00004000
        TransNone    = &H00008000
        TransWall    = &H00010000
        TransGlass   = &H00020000
        TransSteam   = &H00040000
        TransEnergy  = &H00080000
'       Left_Hand    = &H01000000
'       Right_Hand   = &H02000000
'       Worn         = &H04000000
'       HiddenItem   = &H08000000
        WallTransEnd = &H10000000
        LightThru    = &H20000000
        Seen         = &H40000000
        ShootThru    = &H80000000
    End Enum

    Enum FlagsExt As Integer
        BigGun     = &H00000100 ' оружие относится к классу Big Guns
        TwoHand    = &H00000200 ' оружие относится к классу двуручных
        Energy     = &H00000400 ' оружие относится к классу Энергетического (sfall)

        Use        = &H00000800 ' объект можно использовать
        UseOn      = &H00001000
        Look       = &H00002000 ' объект можно осмотреть
        Talk       = &H00004000 ' с объектом можно поговорить
        PickUp     = &H00008000 ' объект можно поднять

        Unknown    = &H00800000
        HiddenItem = &H08000000
    End Enum

    Enum Gender As Integer
        Female = 0
        Male   = 1
    End Enum

End Class
