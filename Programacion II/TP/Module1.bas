Attribute VB_Name = "Module1"
Type usuario
    usuario As String
    apellido As String
    pass As String
    localidad As String
End Type

Type localidad
    localidad As String
    codpostal As String
    codlocalidad As String
End Type

Type categoria
    categoria As String
    codcategoria As String
End Type

Type producto
    producto As String
    codproducto As String
    precio As String
    stock As String
    marca As String
End Type

Type compra
    usuario As String
    zapallida As String
End Type

