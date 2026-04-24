VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColumnInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSchemaName As String
Private mTableName As String
Private mColumnName As String
Private mDataType As String
Private mIsNullable As Boolean
Private mIsIdentity As Boolean
Private mIsComputed As Boolean
Private mHasDefault As Boolean
Private mIsRequired As Boolean
Private mOrdinalPosition As Long
Private mCharacterMaxLength As Long
Private mNumericPrecision As Long
Private mNumericScale As Long

Public Property Get SchemaName() As String
    SchemaName = mSchemaName
End Property

Public Property Let SchemaName(ByVal Value As String)
    mSchemaName = Trim$(Value)
End Property

Public Property Get TableName() As String
    TableName = mTableName
End Property

Public Property Let TableName(ByVal Value As String)
    mTableName = Trim$(Value)
End Property

Public Property Get ColumnName() As String
    ColumnName = mColumnName
End Property

Public Property Let ColumnName(ByVal Value As String)
    mColumnName = Trim$(Value)
End Property

Public Property Get DataType() As String
    DataType = mDataType
End Property

Public Property Let DataType(ByVal Value As String)
    mDataType = Trim$(Value)
End Property

Public Property Get IsNullable() As Boolean
    IsNullable = mIsNullable
End Property

Public Property Let IsNullable(ByVal Value As Boolean)
    mIsNullable = Value
End Property

Public Property Get IsIdentity() As Boolean
    IsIdentity = mIsIdentity
End Property

Public Property Let IsIdentity(ByVal Value As Boolean)
    mIsIdentity = Value
End Property

Public Property Get IsComputed() As Boolean
    IsComputed = mIsComputed
End Property

Public Property Let IsComputed(ByVal Value As Boolean)
    mIsComputed = Value
End Property

Public Property Get HasDefault() As Boolean
    HasDefault = mHasDefault
End Property

Public Property Let HasDefault(ByVal Value As Boolean)
    mHasDefault = Value
End Property

Public Property Get IsRequired() As Boolean
    IsRequired = mIsRequired
End Property

Public Property Let IsRequired(ByVal Value As Boolean)
    mIsRequired = Value
End Property

Public Property Get OrdinalPosition() As Long
    OrdinalPosition = mOrdinalPosition
End Property

Public Property Let OrdinalPosition(ByVal Value As Long)
    mOrdinalPosition = Value
End Property

Public Property Get CharacterMaxLength() As Long
    CharacterMaxLength = mCharacterMaxLength
End Property

Public Property Let CharacterMaxLength(ByVal Value As Long)
    mCharacterMaxLength = Value
End Property

Public Property Get NumericPrecision() As Long
    NumericPrecision = mNumericPrecision
End Property

Public Property Let NumericPrecision(ByVal Value As Long)
    mNumericPrecision = Value
End Property

Public Property Get NumericScale() As Long
    NumericScale = mNumericScale
End Property

Public Property Let NumericScale(ByVal Value As Long)
    mNumericScale = Value
End Property
