Imports System.IO

Public Class CritterObj

    Private Const ProtoSize As Integer = Prototypes.ProtoMemberCount.Critter * 4

    Friend proto As Prototypes.CritterProto
    Private proFile As String

    Public Property PrototypeFile As String
        Set(value As String)
            proFile = value
        End Set
        Get
            Return proFile
        End Get
    End Property

    Sub New(proFile As String)
        MyClass.proFile = proFile
    End Sub

    Public Function Load() As Boolean
        Dim isFallout1 = False

        Dim data(ProtoSize - 1) As Byte

        Dim streamFile = File.Open(proFile, FileMode.Open, FileAccess.Read, FileShare.Read)
        If (streamFile.Length = 412) Then
            streamFile.Read(data, 0, ProtoSize - 4)
            isFallout1 = True
        ElseIf (streamFile.Length = 416) Then
            streamFile.Read(data, 0, ProtoSize)
        Else
            streamFile.Close()
            Return False
        End If

        streamFile.Close()
        ProFiles.ReverseLoadData(data, proto)

        If (isFallout1) Then proto.DamageType = 7 ' устанавливаем для указания прототипа формата F1

        Return True
    End Function

    Public Sub Save(savePath As String)
        Dim size = If(proto.DamageType = 7, ProtoSize - 4, ProtoSize) ' 412

        Dim streamFile = File.Open(savePath, FileMode.Create, FileAccess.Write, FileShare.None)
        streamFile.Write(ProFiles.SaveDataReverse(proto), 0, size)
        streamFile.Close()
    End Sub

    Public Function GetStat(statID As Integer) As Integer
        Select Case statID
            Case 0
                Return proto.Strength
            Case 1
                Return proto.Perception
            Case 2
                Return proto.Endurance
            Case 3
                Return proto.Charisma
            Case 4
                Return proto.Intelligence
            Case 5
                Return proto.Agility
            Case 6
                Return proto.Luck
        End Select
        Return 0
    End Function

#Region "Prototype propertes"

    ''' <summary>
    '''  Устанавливает или возвращает интенсивность в процентах 0..100
    ''' </summary>
    Public Property LightIntensity As Integer
        Get
            Return CInt(Math.Round((proto.LightInt * 100) / 65535))
        End Get
        Set(value As Integer)
            proto.LightInt = CInt(Math.Round((value * 65535) / 100))
        End Set
    End Property

#Region "Flags"

    Public Property IsFlat As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.Flat, proto.Flags And Not (Enums.Flags.Flat))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.Flat) <> 0
        End Get
    End Property

    Public Property IsNoBlock As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.NoBlock, proto.Flags And Not (Enums.Flags.NoBlock))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.NoBlock) <> 0
        End Get
    End Property

    Public Property IsMultiHex As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.MultiHex, proto.Flags And Not (Enums.Flags.MultiHex))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.MultiHex) <> 0
        End Get
    End Property

    Public Property IsShootThru As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.ShootThru, proto.Flags And Not (Enums.Flags.ShootThru))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.ShootThru) <> 0
        End Get
    End Property

    Public Property IsLightThru As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.LightThru, proto.Flags And Not (Enums.Flags.LightThru))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.LightThru) <> 0
        End Get
    End Property

    Public ReadOnly Property IsLighting As Boolean
        Get
            Return (proto.Flags And Enums.Flags.Lighting) <> 0
        End Get
    End Property

    Public Property IsTransNone As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.TransNone, proto.Flags And Not (Enums.Flags.TransNone))
            If (value) Then proto.Flags = proto.Flags And &HFFF0BFFF ' сбросить
        End Set
        Get
            Return (proto.Flags And Enums.Flags.TransNone) <> 0
        End Get
    End Property

    Public Property IsTransWall As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.TransWall, proto.Flags And Not (Enums.Flags.TransWall))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.TransWall) <> 0
        End Get
    End Property

    Public Property IsTransGlass As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.TransGlass, proto.Flags And Not (Enums.Flags.TransGlass))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.TransGlass) <> 0
        End Get
    End Property

    Public Property IsTransSteam As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.TransSteam, proto.Flags And Not (Enums.Flags.TransSteam))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.TransSteam) <> 0
        End Get
    End Property

    Public Property IsTransEnergy As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.TransEnergy, proto.Flags And Not (Enums.Flags.TransEnergy))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.TransEnergy) <> 0
        End Get
    End Property

    Public Property IsTransRed As Boolean
        Set(value As Boolean)
            proto.Flags = If(value, proto.Flags Or Enums.Flags.TransRed, proto.Flags And Not (Enums.Flags.TransRed))
        End Set
        Get
            Return (proto.Flags And Enums.Flags.TransRed) <> 0
        End Get
    End Property

#End Region

#Region "Flags Ext"

    Public Property IsLook As Boolean
        Set(value As Boolean)
            proto.FlagsExt = If(value, proto.Flags Or Enums.FlagsExt.Look, proto.Flags And Not (Enums.FlagsExt.Look))
        End Set
        Get
            Return (proto.FlagsExt And Enums.FlagsExt.Look) <> 0
        End Get
    End Property

    Public Property IsTalk As Boolean
        Set(value As Boolean)
            proto.FlagsExt = If(value, proto.Flags Or Enums.FlagsExt.Talk, proto.Flags And Not (Enums.FlagsExt.Talk))
        End Set
        Get
            Return (proto.FlagsExt And Enums.FlagsExt.Talk) <> 0
        End Get
    End Property

#End Region

#Region "Critter Flags"

    Public Property IsBarter As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.Barter, proto.CritterFlags And Not (Enums.CritterFlags.Barter))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.Barter) <> 0
        End Get
    End Property

    Public Property IsNoSteal As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.NoSteal, proto.CritterFlags And Not (Enums.CritterFlags.NoSteal))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.NoSteal) <> 0
        End Get
    End Property

    Public Property IsNoDrop As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.NoDrop, proto.CritterFlags And Not (Enums.CritterFlags.NoDrop))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.NoDrop) <> 0
        End Get
    End Property


    Public Property IsNoLimbs As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.NoLimbs, proto.CritterFlags And Not (Enums.CritterFlags.NoLimbs))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.NoLimbs) <> 0
        End Get
    End Property

    Public Property IsAges As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.Ages, proto.CritterFlags And Not (Enums.CritterFlags.Ages))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.Ages) <> 0
        End Get
    End Property

    Public Property IsNoHeal As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.NoHeal, proto.CritterFlags And Not (Enums.CritterFlags.NoHeal))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.NoHeal) <> 0
        End Get
    End Property

    Public Property IsInvulnerable As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.Invulnerable, proto.CritterFlags And Not (Enums.CritterFlags.Invulnerable))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.Invulnerable) <> 0
        End Get
    End Property

    Public Property IsNoFlatten As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.NoFlatten, proto.CritterFlags And Not (Enums.CritterFlags.NoFlatten))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.NoFlatten) <> 0
        End Get
    End Property

    Public Property IsSpecialDeath As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.SpecialDeath, proto.CritterFlags And Not (Enums.CritterFlags.SpecialDeath))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.SpecialDeath) <> 0
        End Get
    End Property

    Public Property IsRangeHtH As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.RangeHtH, proto.CritterFlags And Not (Enums.CritterFlags.RangeHtH))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.RangeHtH) <> 0
        End Get
    End Property

    Public Property IsNoKnockBack As Boolean
        Set(value As Boolean)
            proto.CritterFlags = If(value, proto.CritterFlags Or Enums.CritterFlags.NoKnockBack, proto.CritterFlags And Not (Enums.CritterFlags.NoKnockBack))
        End Set
        Get
            Return (proto.CritterFlags And Enums.CritterFlags.NoKnockBack) <> 0
        End Get
    End Property

#End Region

    Public Property ProtoID As Integer
        Set(value As Integer)
            proto.ProtoID = value
        End Set
        Get
            Return proto.ProtoID
        End Get
    End Property

    Public Property DescID As Integer
        Set(value As Integer)
            proto.DescID = value
        End Set
        Get
            Return proto.DescID
        End Get
    End Property

    Public Property FrmID As Integer
        Set(value As Integer)
            proto.FrmID = value
        End Set
        Get
            Return proto.FrmID
        End Get
    End Property

    Public Property LightDis As Integer
        Set(value As Integer)
            proto.LightDis = value
        End Set
        Get
            Return proto.LightDis
        End Get
    End Property

    Public Property LightInt As Integer
        Set(value As Integer)
            proto.LightInt = value
        End Set
        Get
            Return proto.LightInt
        End Get
    End Property

    Public Property Flags As Integer
        Set(value As Integer)
            proto.Flags = value
        End Set
        Get
            Return proto.Flags
        End Get
    End Property

    Public Property FlagsExt As Integer
        Set(value As Integer)
            proto.FlagsExt = value
        End Set
        Get
            Return proto.FlagsExt
        End Get
    End Property

    Public Property ScriptID As Integer
        Set(value As Integer)
            proto.ScriptID = value
        End Set
        Get
            Return proto.ScriptID
        End Get
    End Property

    Public Property HeadFID As Integer
        Set(value As Integer)
            proto.HeadFID = value
        End Set
        Get
            Return proto.HeadFID
        End Get
    End Property

    Public Property AIPacket As Integer
        Set(value As Integer)
            proto.AIPacket = value
        End Set
        Get
            Return proto.AIPacket
        End Get
    End Property

    Public Property TeamNum As Integer
        Set(value As Integer)
            proto.TeamNum = value
        End Set
        Get
            Return proto.TeamNum
        End Get
    End Property

    Public Property CritterFlags As Integer
        Set(value As Integer)
            proto.CritterFlags = value
        End Set
        Get
            Return proto.CritterFlags
        End Get
    End Property

#End Region

End Class
