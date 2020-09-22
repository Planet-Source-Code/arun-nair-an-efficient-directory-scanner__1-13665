VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScanDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private originalPath ' the original path that was passed
Private dirCtr
Private arrCtr
Private dirPathArr() 'array to be dynamically dimensioned

Private Function ScanDir(ByRef sPath, Optional ByRef oPath, Optional ByRef default As Boolean) As Boolean
    
    Dim oldPath, scanPath, parseCtr, ctr
    Dim defVal As Boolean
    
    
    'Change the directory Path to the scan path
    Form1.Dir1.Path = sPath
    
    'if the dir counter is greater than 1
    ' and the oldpath is the original path, STOP
    If dirCtr > 1 And oPath = originalPath Then
        Exit Function
    End If
    
    oldPath = oPath 'the old path for recursion
    
    'initialize the counter
    ctr = 0
    
    'parse throught the list to search for old entries
    For parseCtr = 0 To Form1.Dir1.ListCount - 1
        
        'continue process events as the parsing happens
        DoEvents
        
        ' if the ctr=1 then the old path was found
        If ctr = 1 Then
            scanPath = Form1.Dir1.List(parseCtr)
            Exit For
        End If
        
        'the old path has been found increment the counter
        ' by 1
        If Form1.Dir1.List(parseCtr) = oldPath Then ctr = ctr + 1
        
    Next
    
    
    'the new oldPath which will be sent for recursion
    oldPath = scanPath
    
    
    
    
    'if the path to be send for scanning after the
    'parsing for already checked entries is empty ...
    
    If Trim(scanPath) = "" Then
        'ctr=1 viz. the last directory in the list
        'was already scanned. Hence time to go a level
        ' back
        If ctr >= 1 Then
 
            scanPath = ".."
            oldPath = Form1.Dir1.List(Form1.Dir1.ListIndex)
                
        Else
        'ctr=0 which means no path in this directory
        ' has been scanned
            If Form1.Dir1.ListCount = 0 Then
              
                'The directory count =0
                ' there is nothing in this directory
                
                
                'if this value = false then
                ' the array or listbox must
                ' be populated because this directory
                ' is not the default one(since the
                ' default is applied only to dirs that
                ' have atleast one sub-directory)
                
                scanPath = ".."
                oldPath = Form1.Dir1.List(Form1.Dir1.ListIndex)
                
                If default = False Then
                    
                    dirCtr = dirCtr + 1
                            
                    arrCtr = arrCtr + 1
                    ReDim Preserve dirPathArr(arrCtr)
                    dirPathArr(arrCtr) = oldPath
                    
                End If
                
                'go back one level since the dir is empty
                
            ElseIf Form1.Dir1.ListCount > 0 Then
                'since no dirs have been scanned in this
                'list, default to the first one directory
                ' in the list
                
                 '= form1.dir1.List(0)
                oldPath = Form1.Dir1.List(0)
                scanPath = oldPath
                
                defVal = True
                
                'increment the value of the array counter
                arrCtr = arrCtr + 1
                'preserve the array contents
                ReDim Preserve dirPathArr(arrCtr)
                dirPathArr(arrCtr) = oldPath
          
                dirCtr = dirCtr + 1
                'Call scanDir(form1.dir1.List(0), , True)
                
            End If
        End If
    End If
    
    Call ScanDir(scanPath, oldPath, defVal)
    
End Function

Public Function entryPoint(ByRef getPath) As Variant


On Error GoTo chkerr
    'initialize the dir counter to 1
    dirCtr = 1
    
    'erase the contents of the array
    Erase dirPathArr
    'initialize the array counter
    arrCtr = 0
    
    originalPath = getPath
    Call ScanDir(originalPath, "")
    
    'return the array contents
    entryPoint = dirPathArr
    Exit Function

chkerr:
    
    ' if path not found then stop the scan
    If Err.Number = 76 Then
        
        'return the array contents
        entryPoint = dirPathArr
    ElseIf Err.Number = 28 Then
        'Out-of-stack error
        'pass the last path before the error occurred
        ' for recursion
        
        Call ScanDir(Form1.Dir1.Path, "")
    Else
        Err.Raise vbObjectError + 700, "ArunSysUtils.ScanDir", Err.Description
    End If

End Function
    

