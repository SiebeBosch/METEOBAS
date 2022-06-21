Public Class clsUnitConversion

  Public Function m3ps2mmph(ByVal Value As Double, ByVal AreaM2 As Double) As Double
    If AreaM2 > 0 Then
      Return Value / AreaM2 * 1000 * 3600
    Else
      Return 0
    End If
  End Function

  Public Function mmph2m3ps(ByVal Value As Double, ByVal AreaM2 As Double) As Double
    If AreaM2 > 0 Then
      Return Value / 1000 * AreaM2 / 3600
    Else
      Return 0
    End If
  End Function


End Class
