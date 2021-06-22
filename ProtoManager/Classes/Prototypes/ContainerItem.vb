Imports System.IO
Imports System.Runtime.InteropServices

Public Class ContainerItemObj
    Inherits ItemPrototype

    Private ReadOnly Property ProtoSize As Integer = 2 * 4

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure ContainerProto
        Friend MaxSize As Integer
        Friend Flags As Integer
    End Structure

    Private mProto As ContainerProto

    Sub New(data As ItemPrototype)
        MyBase.New(data)

        ObjType = Enums.ItemType.Container
    End Sub

    Sub New(proFile As String)
        MyBase.New(proFile)
    End Sub

    Public Overloads Sub Load()
        Dim streamFile As BinaryReader = New BinaryReader(MyBase.DataLoad())

        mProto.MaxSize = ProFiles.ReverseBytes(streamFile.ReadInt32)
        mProto.Flags = ProFiles.ReverseBytes(streamFile.ReadInt32)

        streamFile.Close()
    End Sub

    Public Overloads Sub Save(savePath As String)
        Dim streamFile As BinaryWriter = New BinaryWriter(MyBase.DataSave(savePath))

        streamFile.Write(ProFiles.ReverseBytes(mProto.MaxSize))
        streamFile.Write(ProFiles.ReverseBytes(mProto.Flags))

        streamFile.Close()
    End Sub

#Region "Prototype propertes"

    Public WriteOnly Property SetOpenFlag As Boolean
        Set(value As Boolean)
            OpenFlags = If(value, 1, 0)
        End Set
    End Property

    Public ReadOnly Property GetOpenFlag As Boolean
        Get
            Return (OpenFlags And &H1) <> 0
        End Get
    End Property

    Public Property MaxSize As Integer
        Set(value As Integer)
            mProto.MaxSize = value
        End Set
        Get
            Return mProto.MaxSize
        End Get
    End Property

    Public Property OpenFlags As Integer
        Set(value As Integer)
            mProto.Flags = value
        End Set
        Get
            Return mProto.Flags
        End Get
    End Property

#End Region

End Class
