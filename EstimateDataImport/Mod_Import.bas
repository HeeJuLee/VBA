Attribute VB_Name = "Mod_Import"
Option Explicit

Sub ImportAll()

    Dim startTime As Single
    Dim endTime As Single
    
    startTime = Timer
    ImportManage
    endTime = Timer
    shtResult.Range("A2").Value = "���� ó�� �ð� (��) "
    shtResult.Range("B2").Value = Format(endTime - startTime, "#0.00")
    
    startTime = Timer
    ImportEstimate
    endTime = Timer
    shtResult.Range("A3").Value = "�������� ó�� �ð� (��) "
    shtResult.Range("B3").Value = Format(endTime - startTime, "#0.00")
    
    startTime = Timer
    JoinEstimateAccepted
    endTime = Timer
    shtResult.Range("A4").Value = "�������� ���� ó�� �ð� (��) "
    shtResult.Range("B4").Value = Format(endTime - startTime, "#0.00")
    
    startTime = Timer
    JoinOrderEstimate
    endTime = Timer
    shtResult.Range("A5").Value = "���ְ��� ���� ó�� �ð� (��) "
    shtResult.Range("B5").Value = Format(endTime - startTime, "#0.00")
      
    
End Sub
