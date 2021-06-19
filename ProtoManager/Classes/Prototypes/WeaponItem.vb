﻿Imports System.IO
Imports System.Runtime.InteropServices

Public Class WeaponItemObj
    Inherits ItemPrototype

    Enum WeaponType As Integer
        Big = Enums.FlagsExt.BigGun
        TwoHand = Enums.FlagsExt.TwoHand
        Energy = Enums.FlagsExt.Energy
    End Enum

    Enum DamageType As Integer
        Normal
        Laser
        Fire
        Plasma
        Electrical
        EMP
        Explode
    End Enum

    Private ReadOnly Property ProtoSize As Integer = 16 * 4

    <StructLayout(LayoutKind.Sequential, Pack := 1)>
    Public Structure WeaponProto
        Friend AnimCode As Integer
        Friend MinDmg As Integer
        Friend MaxDmg As Integer
        Friend DmgType As Integer
        Friend MaxRangeP As Integer
        Friend MaxRangeS As Integer
        Friend ProjPID As Integer
        Friend MinST As Integer
        Friend MPCostP As Integer
        Friend MPCostS As Integer
        Friend CritFail As Integer
        Friend Perk As Integer
        Friend Rounds As Integer
        Friend Caliber As Integer
        Friend AmmoPID As Integer
        Friend MaxAmmo As Integer
        Friend wSoundID As Byte
    End Structure

    Private mProto As WeaponProto

    Sub New(proFile As String)
        MyBase.New(proFile)
    End Sub

    Public Overloads Sub Load()
        Dim streamFile = MyBase.DataLoad()

        Dim data(ProtoSize - 1) As Byte
        streamFile.Read(data, 0, ProtoSize)

        ProFiles.ReverseLoadData(data, mProto)
        mProto.wSoundID = CByte(streamFile.ReadByte())
        streamFile.Close()
    End Sub

    Public Overloads Sub Save(savePath As String)
        Dim streamFile = MyBase.DataSave(savePath)

        streamFile.Write(ProFiles.SaveDataReverse(mProto), 0, ProtoSize + 1)
        streamFile.Close()
    End Sub

    Public Function WeaponScore() As Integer
        Dim Score As Integer = CInt(Math.Floor(Math.Abs(mProto.MaxDmg - mProto.MinDmg) / 2))
        If (mProto.Perk > 0) Then Score *= 5
        Return Score
    End Function

#Region "Prototype propertes"

    Public Property PrimaryAttackType As Integer
        Set(value As Integer)
            Dim flags = MyBase.FlagsExt And &HFFFFFFF0
            MyBase.FlagsExt = flags Or (value And &HF)
        End Set
        Get
            Return MyBase.FlagsExt And &HF
        End Get
    End Property

    Public Property SecondaryAttackType As Integer
        Set(value As Integer)
            Dim flags = MyBase.FlagsExt And &HFFFFFF0F
            MyBase.FlagsExt = flags Or ((value And &HF) << 4)
        End Set
        Get
            Return (MyBase.FlagsExt >> 4) And &HF
        End Get
    End Property

    Public Property IsBigGun As Boolean
        Set(value As Boolean)
            MyBase.FlagsExt = If(value, MyBase.FlagsExt Or WeaponType.Big, MyBase.FlagsExt And Not (WeaponType.Big))
        End Set
        Get
            Return (MyBase.FlagsExt And WeaponType.Big) <> 0
        End Get
    End Property

    Public Property IsTwoHand As Boolean
        Set(value As Boolean)
            MyBase.FlagsExt = If(value, MyBase.FlagsExt Or WeaponType.TwoHand, MyBase.FlagsExt And Not (WeaponType.TwoHand))
        End Set
        Get
            Return (MyBase.FlagsExt And WeaponType.TwoHand) <> 0
        End Get
    End Property

    Public Property IsEnergy As Boolean
        Set(value As Boolean)
            MyBase.FlagsExt = If(value, MyBase.FlagsExt Or WeaponType.Energy, MyBase.FlagsExt And Not (WeaponType.Energy))
        End Set
        Get
            Return (MyBase.FlagsExt And WeaponType.Energy) <> 0
        End Get
    End Property

    Public Property AnimCode() As Integer
        Set(value As Integer)
            mProto.AnimCode = value
        End Set
        Get
            Return mProto.AnimCode
        End Get
    End Property

    Public Property MinDmg() As Integer
        Set(value As Integer)
            mProto.MinDmg = value
        End Set
        Get
            Return mProto.MinDmg
        End Get
    End Property

    Public Property MaxDmg() As Integer
        Set(value As Integer)
            mProto.MaxDmg = value
        End Set
        Get
            Return mProto.MaxDmg
        End Get
    End Property

    Public Property DmgType() As DamageType
        Set(value As DamageType)
            mProto.DmgType = value
        End Set
        Get
            Return CType(mProto.DmgType, DamageType)
        End Get
    End Property

    Public Property MaxRangeP() As Integer
        Set(value As Integer)
            mProto.MaxRangeP = value
        End Set
        Get
            Return mProto.MaxRangeP
        End Get
    End Property

    Public Property MaxRangeS() As Integer
        Set(value As Integer)
            mProto.MaxRangeS = value
        End Set
        Get
            Return mProto.MaxRangeS
        End Get
    End Property

    Public Property ProjPID() As Integer
        Set(value As Integer)
            mProto.ProjPID = value
        End Set
        Get
            Return mProto.ProjPID
        End Get
    End Property

    Public Property MinST() As Integer
        Set(value As Integer)
            mProto.MinST = value
        End Set
        Get
            Return mProto.MinST
        End Get
    End Property

    Public Property MPCostP() As Integer
        Set(value As Integer)
            mProto.MPCostP = value
        End Set
        Get
            Return mProto.MPCostP
        End Get
    End Property

    Public Property MPCostS() As Integer
        Set(value As Integer)
            mProto.MPCostS = value
        End Set
        Get
            Return mProto.MPCostS
        End Get
    End Property

    Public Property CritFail() As Integer
        Set(value As Integer)
            mProto.CritFail = value
        End Set
        Get
            Return mProto.CritFail
        End Get
    End Property

    Public Property Perk() As Integer
        Set(value As Integer)
            mProto.Perk = value
        End Set
        Get
            Return mProto.Perk
        End Get
    End Property

    Public Property Rounds() As Integer
        Set(value As Integer)
            mProto.Rounds = value
        End Set
        Get
            Return mProto.Rounds
        End Get
    End Property

    Public Property Caliber() As Integer
        Set(value As Integer)
            mProto.Caliber = value
        End Set
        Get
            Return mProto.Caliber
        End Get
    End Property

    Public Property AmmoPID() As Integer
        Set(value As Integer)
            mProto.AmmoPID = value
        End Set
        Get
            Return mProto.AmmoPID
        End Get
    End Property

    Public Property MaxAmmo() As Integer
        Set(value As Integer)
            mProto.MaxAmmo = value
        End Set
        Get
            Return mProto.MaxAmmo
        End Get
    End Property

    Public Property wSoundID() As Byte
        Set(value As Byte)
            mProto.wSoundID = value
        End Set
        Get
            Return mProto.wSoundID
        End Get
    End Property

#End Region

End Class
