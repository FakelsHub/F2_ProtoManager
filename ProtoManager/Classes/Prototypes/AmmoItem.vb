﻿Imports System.Runtime.InteropServices

Public Class AmmoItemObj
    Inherits ItemPrototype

    Private ReadOnly Property ProtoSize As Integer = 6 * 4

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure AmmoProto
        Friend Caliber As Integer
        Friend Quantity As Integer
        Friend ACAdjust As Integer
        Friend DRAdjust As Integer
        Friend DamMult As Integer
        Friend DamDiv As Integer
    End Structure

    Private mProto As AmmoProto

    Sub New(data As ItemPrototype)
        MyBase.New(data)

        ObjType = Enums.ItemType.Ammo
        DamMult = 1
        DamDiv = 1
    End Sub

    Sub New(proFile As String)
        MyBase.New(proFile)
    End Sub

    Public Overloads Sub Load()
        Dim streamFile = MyBase.DataLoad()

        Dim data(ProtoSize - 1) As Byte

        streamFile.Read(data, 0, ProtoSize)
        streamFile.Close()

        ProFiles.ReverseLoadData(data, mProto)
    End Sub

    Public Overloads Sub Save(savePath As String)
        Dim streamFile = MyBase.DataSave(savePath)

        streamFile.Write(ProFiles.SaveDataReverse(mProto), 0, ProtoSize)
        streamFile.Close()
    End Sub

#Region "Prototype propertes"

    Public Property Caliber As Integer
        Set(value As Integer)
            mProto.Caliber = value
        End Set
        Get
            Return mProto.Caliber
        End Get
    End Property

    Public Property Quantity As Integer
        Set(value As Integer)
            mProto.Quantity = value
        End Set
        Get
            Return mProto.Quantity
        End Get
    End Property

    Public Property ACAdjust As Integer
        Set(value As Integer)
            mProto.ACAdjust = value
        End Set
        Get
            Return mProto.ACAdjust
        End Get
    End Property

    Public Property DRAdjust As Integer
        Set(value As Integer)
            mProto.DRAdjust = value
        End Set
        Get
            Return mProto.DRAdjust
        End Get
    End Property

    Public Property DamMult As Integer
        Set(value As Integer)
            mProto.DamMult = value
        End Set
        Get
            Return mProto.DamMult
        End Get
    End Property

    Public Property DamDiv As Integer
        Set(value As Integer)
            mProto.DamDiv = value
        End Set
        Get
            Return mProto.DamDiv
        End Get
    End Property

#End Region

End Class
