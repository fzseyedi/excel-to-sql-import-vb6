VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMappingItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mExcelColumnName As String
Private mSqlColumnName As String
Private mSqlDataType As String
Private mIsRequired As Boolean
Private mIsMapped As Boolean

Public Property Get ExcelColumnName() As String
    ExcelColumnName = mExcelColumnName
End Property

Public Property Let ExcelColumnName(ByVal Value As String)
    mExcelColumnName = Trim$(Value)
End Property

Public Property Get SqlColumnName() As String
    SqlColumnName = mSqlColumnName
End Property

Public Property Let SqlColumnName(ByVal Value As String)
    mSqlColumnName = Trim$(Value)
End Property

Public Property Get SqlDataType() As String
    SqlDataType = mSqlDataType
End Property

Public Property Let SqlDataType(ByVal Value As String)
    mSqlDataType = Trim$(Value)
End Property

Public Property Get IsRequired() As Boolean
    IsRequired = mIsRequired
End Property

Public Property Let IsRequired(ByVal Value As Boolean)
    mIsRequired = Value
End Property

Public Property Get IsMapped() As Boolean
    IsMapped = mIsMapped
End Property

Public Property Let IsMapped(ByVal Value As Boolean)
    mIsMapped = Value
End Property
