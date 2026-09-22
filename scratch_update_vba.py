import sys, win32com.client, os

tmpl_path = os.path.abspath(r'IQOQDQ\OQ_HMI\Sequence_Checklist_Template.xlsm')

new_vba_code = """Option Explicit

Sub Checkbox_Numbering()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Call ReNumber_Selected_Items(ws)
End Sub

Private Sub ReNumber_Selected_Items(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim chk As CheckBox
    Dim numFound As Long
    Dim cellList() As Long
    Dim valList() As Double
    Dim tempRow As Long, tempVal As Double
    Dim curVal As Variant

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Application.ScreenUpdating = False

    numFound = 0
    For Each chk In ws.CheckBoxes
        i = chk.TopLeftCell.Row
        If i >= 2 And i <= lastRow Then
            If chk.Value = xlOn Then
                curVal = ws.Cells(i, "C").Value
                If Not IsNumeric(curVal) Or IsEmpty(curVal) Or CDbl(curVal) <= 0 Then
                    curVal = 999999
                    ws.Cells(i, "C").Value = curVal
                End If
                numFound = numFound + 1
                ReDim Preserve cellList(1 To numFound)
                ReDim Preserve valList(1 To numFound)
                cellList(numFound) = i
                valList(numFound) = CDbl(curVal)
            Else
                ws.Cells(i, "C").ClearContents
                ws.Range(ws.Cells(i, "A"), ws.Cells(i, "G")).Interior.ColorIndex = xlNone
            End If
        End If
    Next chk

    If numFound > 0 Then
        ' Sort checked rows by sequence value ascending (preserving row index order for ties)
        For i = 1 To numFound - 1
            For j = i + 1 To numFound
                If valList(i) > valList(j) Then
                    tempVal = valList(i): valList(i) = valList(j): valList(j) = tempVal
                    tempRow = cellList(i): cellList(i) = cellList(j): cellList(j) = tempRow
                End If
            Next j
        Next i

        ' Re-assign contiguous sequence 1, 2, 3... N
        For i = 1 To numFound
            ws.Cells(cellList(i), "C").Value = i
            ws.Range(ws.Cells(cellList(i), "A"), ws.Cells(cellList(i), "G")).Interior.Color = RGB(217, 234, 211)
        Next i
    End If

    Application.ScreenUpdating = True
End Sub
"""

xl = win32com.client.Dispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
wb = xl.Workbooks.Open(tmpl_path)

try:
    c = wb.VBProject.VBComponents('MacroModule')
    cm = c.CodeModule
    cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(new_vba_code)
    wb.Save()
    print('Updated VBA Macro in Sequence_Checklist_Template.xlsm successfully!')
except Exception as e:
    print('Error updating VBA:', e)
finally:
    wb.Close(True)
    xl.Quit()
