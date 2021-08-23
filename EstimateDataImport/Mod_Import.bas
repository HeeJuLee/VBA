Attribute VB_Name = "Mod_Import"
Option Explicit

Sub ImportAll()

    Dim startTime As Single
    Dim endTime As Single
    
    startTime = Timer
    ImportManage
    endTime = Timer
    shtResult.Range("A2").Value = "관리 처리 시간 (초) "
    shtResult.Range("B2").Value = Format(endTime - startTime, "#0.00")
    
    startTime = Timer
    ImportEstimate
    endTime = Timer
    shtResult.Range("A3").Value = "견적관리 처리 시간 (초) "
    shtResult.Range("B3").Value = Format(endTime - startTime, "#0.00")
    
    startTime = Timer
    JoinEstimateAccepted
    endTime = Timer
    shtResult.Range("A4").Value = "견적수주 조인 처리 시간 (초) "
    shtResult.Range("B4").Value = Format(endTime - startTime, "#0.00")
    
    startTime = Timer
    JoinOrderEstimate
    endTime = Timer
    shtResult.Range("A5").Value = "발주견적 조인 처리 시간 (초) "
    shtResult.Range("B5").Value = Format(endTime - startTime, "#0.00")
      
    
End Sub
