Imports System.Runtime.InteropServices

Public Class MiscItemObj
    Inherits ItemPrototype

    Private ReadOnly Property ProtoSize As Integer = 3 * 4

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure MiscProto
        Friend PowerPID As Integer
        Friend PowerType As Integer
        Friend Charges As Integer
    End Structure

    Private mProto As MiscProto

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

    Public Property PowerPID As Integer
        Set(value As Integer)
            mProto.PowerPID = value
        End Set
        Get
            Return mProto.PowerPID
        End Get
    End Property

    Public Property PowerType As Integer
        Set(value As Integer)
            mProto.PowerType = value
        End Set
        Get
            Return mProto.PowerType
        End Get
    End Property

    Public Property Charges As Integer
        Set(value As Integer)
            mProto.Charges = value
        End Set
        Get
            Return mProto.Charges
        End Get
    End Property

#End Region

End Class
