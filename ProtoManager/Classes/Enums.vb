Class Enums

    Enum FlagsExt As Integer
        HiddenItem = &H08000000
        Unknown    = &H00800000

        UseOn      = &H00001000
        Look       = &H00002000 ' объект можно осмотреть
        Talk       = &H00004000 ' с объектом можно поговорить
        PickUp     = &H00008000 ' объект можно поднять
        Use        = &H00000800 ' объект можно использовать

        BigGun     = &H00000100 ' оружие относится к классу Big Guns
        TwoHand    = &H00000200 ' оружие относится к классу двуручных
        Energy     = &H00000400 ' оружие относится к классу Энергетического (sfall)
    End Enum

End Class
