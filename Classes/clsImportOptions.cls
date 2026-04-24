VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDeleteExistingData As Boolean
Private mContinueOnTypeError As Boolean
Private mContinueOnDuplicate As Boolean
Private mAutoLoadSavedMapping As Boolean
Private mUseIdentityInsert As Boolean
Private mExcelReadMode As Integer

Public Property Get DeleteExistingData() As Boolean
    DeleteExistingData = mDeleteExistingData
End Property

Public Property Let DeleteExistingData(ByVal Value As Boolean)
    mDeleteExistingData = Value
End Property

Public Property Get ContinueOnTypeError() As Boolean
    ContinueOnTypeError = mContinueOnTypeError
End Property

Public Property Let ContinueOnTypeError(ByVal Value As Boolean)
    mContinueOnTypeError = Value
End Property

Public Property Get ContinueOnDuplicate() As Boolean
    ContinueOnDuplicate = mContinueOnDuplicate
End Property

Public Property Let ContinueOnDuplicate(ByVal Value As Boolean)
    mContinueOnDuplicate = Value
End Property

Public Property Get AutoLoadSavedMapping() As Boolean
    AutoLoadSavedMapping = mAutoLoadSavedMapping
End Property

Public Property Let AutoLoadSavedMapping(ByVal Value As Boolean)
    mAutoLoadSavedMapping = Value
End Property

Public Property Get UseIdentityInsert() As Boolean
    UseIdentityInsert = mUseIdentityInsert
End Property

Public Property Let UseIdentityInsert(ByVal Value As Boolean)
    mUseIdentityInsert = Value
End Property

Public Property Get ExcelReadMode() As Integer
    ExcelReadMode = mExcelReadMode
End Property

Public Property Let ExcelReadMode(ByVal Value As Integer)
    mExcelReadMode = Value
End Property
