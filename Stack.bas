Attribute VB_Name = "Stack"
Dim m_vStack() As Variant
Function Pop() As Variant
    On Error Resume Next
    Dim iM As Integer
    iM = UBound(m_vStack)
    Pop = m_vStack(iM)
    iM = iM - 1
    ReDim Preserve m_vStack(iM) As Variant
End Function
Private Sub Smash()
    ReDim m_vStack(0)
End Sub
Private Function Size() As Integer
    On Error Resume Next
    Size = UBound(m_vStack)
End Function
Sub Push(ByVal vValue As Variant)
    Dim iM As Integer
    On Error Resume Next
    iM = UBound(m_vStack)
    iM = iM + 1
    ReDim Preserve m_vStack(iM) As Variant
    m_vStack(iM) = vValue
End Sub

