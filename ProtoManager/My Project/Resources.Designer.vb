﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Этот код создан программой.
'     Исполняемая версия:4.0.30319.233
'
'     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
'     повторной генерации кода.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'Этот класс создан автоматически классом StronglyTypedResourceBuilder
    'с помощью такого средства, как ResGen или Visual Studio.
    'Чтобы добавить или удалить член, измените файл .ResX и снова запустите ResGen
    'с параметром /str или перестройте свой проект VS.
    '''<summary>
    '''  Класс ресурса со строгой типизацией для поиска локализованных строк и т.д.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Возвращает кэшированный экземпляр ResourceManager, использованный этим классом.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Перезаписывает свойство CurrentUICulture текущего потока для всех
        '''  обращений к ресурсу с помощью этого класса ресурса со строгой типизацией.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        Friend ReadOnly Property ammo() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("ammo", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        Friend ReadOnly Property armor() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("armor", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        Friend ReadOnly Property container() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("container", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        Friend ReadOnly Property critter() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("critter", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        Friend ReadOnly Property drugs() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("drugs", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        Friend ReadOnly Property key() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("key", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        Friend ReadOnly Property LeftArrow() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("LeftArrow", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property misc() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("misc", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
        
        Friend ReadOnly Property RESERVAA() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("RESERVAA", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property RightArrow() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("RightArrow", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property splash() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("splash", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        Friend ReadOnly Property weapon() As Byte()
            Get
                Dim obj As Object = ResourceManager.GetObject("weapon", resourceCulture)
                Return CType(obj,Byte())
            End Get
        End Property
    End Module
End Namespace
