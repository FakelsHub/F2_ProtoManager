Imports System.IO
Imports System.Runtime.InteropServices

Public Class DrugsItemObj
    Inherits ItemPrototype

    Private ReadOnly Property ProtoSize As Integer = 17 * 4

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure DrugsProto
        Friend Stat0 As Integer
        Friend Stat1 As Integer
        Friend Stat2 As Integer
        Friend iAmount0 As Integer
        Friend iAmount1 As Integer
        Friend iAmount2 As Integer
        Friend Duration1 As Integer
        Friend fAmount0 As Integer
        Friend fAmount1 As Integer
        Friend fAmount2 As Integer
        Friend Duration2 As Integer
        Friend sAmount0 As Integer
        Friend sAmount1 As Integer
        Friend sAmount2 As Integer
        Friend AddictionRate As Integer
        Friend W_Effect As Integer
        Friend W_Onset As Integer
    End Structure

    Private mProto As DrugsProto

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

    Public Property Stat0 As Integer
        Set(value As Integer)
            mProto.Stat0 = value
        End Set
        Get
            Return mProto.Stat0
        End Get
    End Property

    Public Property Stat1 As Integer
        Set(value As Integer)
            mProto.Stat1 = value
        End Set
        Get
            Return mProto.Stat1
        End Get
    End Property

    Public Property Stat2 As Integer
        Set(value As Integer)
            mProto.Stat2 = value
        End Set
        Get
            Return mProto.Stat2
        End Get
    End Property

    Public Property iAmount0 As Integer
        Set(value As Integer)
            mProto.iAmount0 = value
        End Set
        Get
            Return mProto.iAmount0
        End Get
    End Property

    Public Property iAmount1 As Integer
        Set(value As Integer)
            mProto.iAmount1 = value
        End Set
        Get
            Return mProto.iAmount1
        End Get
    End Property

    Public Property iAmount2 As Integer
        Set(value As Integer)
            mProto.iAmount2 = value
        End Set
        Get
            Return mProto.iAmount2
        End Get
    End Property

    Public Property Duration1 As Integer
        Set(value As Integer)
            mProto.Duration1 = value
        End Set
        Get
            Return mProto.Duration1
        End Get
    End Property

    Public Property fAmount0 As Integer
        Set(value As Integer)
            mProto.fAmount0 = value
        End Set
        Get
            Return mProto.fAmount0
        End Get
    End Property

    Public Property fAmount1 As Integer
        Set(value As Integer)
            mProto.fAmount1 = value
        End Set
        Get
            Return mProto.fAmount1
        End Get
    End Property

    Public Property fAmount2 As Integer
        Set(value As Integer)
            mProto.fAmount2 = value
        End Set
        Get
            Return mProto.fAmount2
        End Get
    End Property

    Public Property Duration2 As Integer
        Set(value As Integer)
            mProto.Duration2 = value
        End Set
        Get
            Return mProto.Duration2
        End Get
    End Property

    Public Property sAmount0 As Integer
        Set(value As Integer)
            mProto.sAmount0 = value
        End Set
        Get
            Return mProto.sAmount0
        End Get
    End Property

    Public Property sAmount1 As Integer
        Set(value As Integer)
            mProto.sAmount1 = value
        End Set
        Get
            Return mProto.sAmount1
        End Get
    End Property

    Public Property sAmount2 As Integer
        Set(value As Integer)
            mProto.sAmount2 = value
        End Set
        Get
            Return mProto.sAmount2
        End Get
    End Property

    Public Property AddictionRate As Integer
        Set(value As Integer)
            mProto.AddictionRate = value
        End Set
        Get
            Return mProto.AddictionRate
        End Get
    End Property

    Public Property W_Effect As Integer
        Set(value As Integer)
            mProto.W_Effect = value
        End Set
        Get
            Return mProto.W_Effect
        End Get
    End Property

    Public Property W_Onset As Integer
        Set(value As Integer)
            mProto.W_Onset = value
        End Set
        Get
            Return mProto.W_Onset
        End Get
    End Property

#End Region

End Class
