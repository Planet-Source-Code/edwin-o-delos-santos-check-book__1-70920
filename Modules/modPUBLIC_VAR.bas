Attribute VB_Name = "modPUBLIC_VAR"
Option Explicit
Public nxTab     As Integer     'TO HANDLE TAB ORDER/keyEvents [Enter] & Up arrow key
Public entries() As TextBox   'dynamic array for entries
Public item_added() As Variant  'used by add_item, already_added proc
Public addRec As Boolean        'to handle new record
Public editRec As Boolean       'to handle existing record
Public currLen(15) As Integer   'array to store lenght of string/value - used by PrintValue procedure
Public printIndex(50) As String 'store field to print
Public initPrint As Boolean     'initialize list
'** var to handle to move form
Public down As Boolean
Public t As Integer
Public w As Integer
'**--------------------------
