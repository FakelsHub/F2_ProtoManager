Imports System.IO

Public Class KeyItemObj
    Inherits ItemPrototype

    Private ReadOnly Property ProtoSize As Integer = 1 * 4

    Public Structure KeyProto
        Friend Unknown As Integer
    End Structure

    Private mProto As KeyProto

    Sub New(proFile As String)
        MyBase.New(proFile)
    End Sub

    Public Overloads Sub Load()
        Dim streamFile As BinaryReader = New BinaryReader(MyBase.DataLoad())

        mProto.Unknown = ProFiles.ReverseBytes(streamFile.ReadInt32)
        streamFile.Close()
    End Sub

    Public Overloads Sub Save(savePath As String)
        Dim streamFile As BinaryWriter = New BinaryWriter(MyBase.DataSave(savePath))
        streamFile.Write(ProFiles.ReverseBytes(mProto.Unknown))
        streamFile.Close()
    End Sub

#Region "Prototype propertes"

    Public Property Unknown As Integer
        Set(value As Integer)
            mProto.Unknown = value
        End Set
        Get
            Return mProto.Unknown
        End Get
    End Property

#End Region
End Class
