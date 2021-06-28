Imports System.IO
Imports System.Runtime.InteropServices

Public Class ItemPrototype
    Implements IPrototype

    Private ReadOnly Property CommonSize As Integer = 14 * 4
    'Friend Property ProtoIsLoad As Boolean = False

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Structure CommonItemProto
        Friend ProtoID As Integer
        Friend DescID As Integer
        Friend FrmID As Integer
        Friend LightDis As Integer
        Friend LightInt As Integer
        Friend Flags As Integer
        Friend FlagsExt As Integer
        Friend ScriptID As Integer
        Friend ObjType As Integer
        Friend MaterialID As Integer
        Friend Size As Integer
        Friend Weight As Integer
        Friend Cost As Integer
        Friend InvFID As Integer
        Friend SoundID As Byte
    End Structure

    Private mProto As CommonItemProto
    Private proFile As String

    Public Property PrototypeFile As String
        Set(value As String)
            proFile = value
        End Set
        Get
            Return proFile
        End Get
    End Property

    Sub New(data As ItemPrototype)
        mProto = data.mProto
        proFile = data.proFile

        FlagsExt = FlagsExt And &HCFFFFF ' default set
    End Sub

    Sub New(proFile As String)
        MyClass.proFile = proFile
    End Sub

    Public Sub Load() Implements IPrototype.Load
        DataLoad().Close()
        'ProtoIsLoad = True
    End Sub

    Public Sub Save(savePath As String) Implements IPrototype.Save
        DataSave(savePath).Close()
    End Sub

    Protected Function DataLoad() As FileStream
        Dim streamFile = File.Open(proFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

        Dim data(CommonSize - 1) As Byte
        streamFile.Read(data, 0, CommonSize)

        ProFiles.ReverseLoadData(data, mProto)
        mProto.SoundID = CByte(streamFile.ReadByte())

        Return streamFile
    End Function

    Protected Function DataSave(savePath As String) As FileStream
        Dim streamFile = File.Open(savePath, FileMode.Create, FileAccess.Write, FileShare.None)

        streamFile.Write(ProFiles.SaveDataReverse(mProto), 0, CommonSize + 1)
        Return streamFile
    End Function

#Region "Prototype propertes"

#Region "Flags"
    Public Property IsFlat As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.Flat, mProto.Flags And Not (Enums.Flags.Flat))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.Flat) <> 0
        End Get
    End Property

    Public Property IsNoBlock As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.NoBlock, mProto.Flags And Not (Enums.Flags.NoBlock))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.NoBlock) <> 0
        End Get
    End Property

    Public Property IsMultiHex As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.MultiHex, mProto.Flags And Not (Enums.Flags.MultiHex))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.MultiHex) <> 0
        End Get
    End Property

    Public Property IsShootThru As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.ShootThru, mProto.Flags And Not (Enums.Flags.ShootThru))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.ShootThru) <> 0
        End Get
    End Property

    Public Property IsLightThru As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.LightThru, mProto.Flags And Not (Enums.Flags.LightThru))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.LightThru) <> 0
        End Get
    End Property

    Public ReadOnly Property IsLighting As Boolean
        Get
            Return (mProto.Flags And Enums.Flags.Lighting) <> 0
        End Get
    End Property

    Public Property IsNoHighlight As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.NoHighlight, mProto.Flags And Not (Enums.Flags.NoHighlight))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.NoHighlight) <> 0
        End Get
    End Property

    Public Property IsTransNone As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.TransNone, mProto.Flags And Not (Enums.Flags.TransNone))
            If (value) Then mProto.Flags = mProto.Flags And &HFFF0BFFF ' сбросить
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.TransNone) <> 0
        End Get
    End Property

    Public Property IsTransWall As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.TransWall, mProto.Flags And Not (Enums.Flags.TransWall))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.TransWall) <> 0
        End Get
    End Property

    Public Property IsTransGlass As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.TransGlass, mProto.Flags And Not (Enums.Flags.TransGlass))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.TransGlass) <> 0
        End Get
    End Property

    Public Property IsTransSteam As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.TransSteam, mProto.Flags And Not (Enums.Flags.TransSteam))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.TransSteam) <> 0
        End Get
    End Property

    Public Property IsTransEnergy As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.TransEnergy, mProto.Flags And Not (Enums.Flags.TransEnergy))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.TransEnergy) <> 0
        End Get
    End Property

    Public Property IsTransRed As Boolean
        Set(value As Boolean)
            mProto.Flags = If(value, mProto.Flags Or Enums.Flags.TransRed, mProto.Flags And Not (Enums.Flags.TransRed))
        End Set
        Get
            Return (mProto.Flags And Enums.Flags.TransRed) <> 0
        End Get
    End Property
#End Region

#Region "Flags Ext"
    Public Property IsUse As Boolean
        Set(value As Boolean)
            mProto.FlagsExt = If(value, mProto.Flags Or Enums.FlagsExt.Use, mProto.Flags And Not (Enums.FlagsExt.Use))
        End Set
        Get
            Return (mProto.FlagsExt And Enums.FlagsExt.Use) <> 0
        End Get
    End Property

    Public Property IsUseOn As Boolean
        Set(value As Boolean)
            mProto.FlagsExt = If(value, mProto.Flags Or Enums.FlagsExt.UseOn, mProto.Flags And Not (Enums.FlagsExt.UseOn))
        End Set
        Get
            Return (mProto.FlagsExt And Enums.FlagsExt.UseOn) <> 0
        End Get
    End Property

    Public Property IsLook As Boolean
        Set(value As Boolean)
            mProto.FlagsExt = If(value, mProto.Flags Or Enums.FlagsExt.Look, mProto.Flags And Not (Enums.FlagsExt.Look))
        End Set
        Get
            Return (mProto.FlagsExt And Enums.FlagsExt.Look) <> 0
        End Get
    End Property

    Public Property IsPickUp As Boolean
        Set(value As Boolean)
            mProto.FlagsExt = If(value, mProto.Flags Or Enums.FlagsExt.PickUp, mProto.Flags And Not (Enums.FlagsExt.PickUp))
        End Set
        Get
            Return (mProto.FlagsExt And Enums.FlagsExt.PickUp) <> 0
        End Get
    End Property

    Public Property IsTalk As Boolean
        Set(value As Boolean)
            mProto.FlagsExt = If(value, mProto.Flags Or Enums.FlagsExt.Talk, mProto.Flags And Not (Enums.FlagsExt.Talk))
        End Set
        Get
            Return (mProto.FlagsExt And Enums.FlagsExt.Talk) <> 0
        End Get
    End Property

    Public Property IsHiddenItem As Boolean
        Set(value As Boolean)
            mProto.FlagsExt = If(value, mProto.Flags Or Enums.FlagsExt.HiddenItem, mProto.Flags And Not (Enums.FlagsExt.HiddenItem))
        End Set
        Get
            Return (mProto.FlagsExt And Enums.FlagsExt.HiddenItem) <> 0
        End Get
    End Property
#End Region

    Public Property ProtoID As Integer
        Set(value As Integer)
            mProto.ProtoID = value
        End Set
        Get
            Return mProto.ProtoID
        End Get
    End Property

    Public Property DescID As Integer
        Set(value As Integer)
            mProto.DescID = value
        End Set
        Get
            Return mProto.DescID
        End Get
    End Property

    Public Property FrmID As Integer
        Set(value As Integer)
            mProto.FrmID = value
        End Set
        Get
            Return mProto.FrmID
        End Get
    End Property

    Public Property LightDis As Integer
        Set(value As Integer)
            mProto.LightDis = value
        End Set
        Get
            Return mProto.LightDis
        End Get
    End Property

    Public Property LightInt As Integer
        Set(value As Integer)
            mProto.LightInt = value
        End Set
        Get
            Return mProto.LightInt
        End Get
    End Property

    Public Property Flags As Integer
        Set(value As Integer)
            mProto.Flags = value
        End Set
        Get
            Return mProto.Flags
        End Get
    End Property

    Public Property FlagsExt As Integer
        Set(value As Integer)
            mProto.FlagsExt = value
        End Set
        Get
            Return mProto.FlagsExt
        End Get
    End Property

    Public Property ScriptID As Integer
        Set(value As Integer)
            mProto.ScriptID = value
        End Set
        Get
            Return mProto.ScriptID
        End Get
    End Property

    Public Property ObjType As Integer
        Set(value As Integer)
            mProto.ObjType = value
        End Set
        Get
            Return mProto.ObjType
        End Get
    End Property

    Public Property MaterialID As Integer
        Set(value As Integer)
            mProto.MaterialID = value
        End Set
        Get
            Return mProto.MaterialID
        End Get
    End Property

    Public Property Size As Integer
        Set(value As Integer)
            mProto.Size = value
        End Set
        Get
            Return mProto.Size
        End Get
    End Property

    Public Property Weight As Integer
        Set(value As Integer)
            mProto.Weight = value
        End Set
        Get
            Return mProto.Weight
        End Get
    End Property

    Public Property Cost As Integer
        Set(value As Integer)
            mProto.Cost = value
        End Set
        Get
            Return mProto.Cost
        End Get
    End Property

    Public Property InvFID As Integer
        Set(value As Integer)
            mProto.InvFID = value
        End Set
        Get
            Return mProto.InvFID
        End Get
    End Property

    Public Property SoundID As Byte
        Set(value As Byte)
            mProto.SoundID = value
        End Set
        Get
            Return mProto.SoundID
        End Get
    End Property

#End Region

End Class
