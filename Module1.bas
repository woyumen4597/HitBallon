Attribute VB_Name = "Module1"
Option Explicit
Public Type user
name As String * 4
fs As Integer
End Type
Public Ulen, MinScore As Integer 'user�ĳ��ȣ����а�����һ���ķ���
Public Rank(1 To 5) As user
Public Score As Double
Public t As Double
