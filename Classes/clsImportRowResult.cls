VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportRowResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRowNumber As Long
Private mIsSuccess As Boolean
Private mIsSkipped As Boolean
Private mErrorMessage As String

Public Property Get RowNumber() As Long
    RowNumber = mRowNumber
End Property

Public Property Let RowNumber(ByVal Value As Long)
    mRowNumber = Value
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = mIsSuccess
End Property

Public Property Let IsSuccess(ByVal Value As Boolean)
    mIsSuccess = Value
End Property

Public Property Get IsSkipped() As Boolean
    IsSkipped = mIsSkipped
End Property

Public Property Let IsSkipped(ByVal Value As Boolean)
    mIsSkipped = Value
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = mErrorMessage
End Property

Public Property Let ErrorMessage(ByVal Value As String)
    mErrorMessage = Trim$(Value)
End Property
