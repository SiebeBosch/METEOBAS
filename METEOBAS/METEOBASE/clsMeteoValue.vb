Public Class clsMeteoValue
  Public Datum As Long
  Public Tijd As Long
  Public DateTimeVal As DateTime
  Public ValueObserved As Double  'original in 0.1 mm
  Public ValueCorrected As Double 'converted to mm and replaced -1 by 0.025mm
  Public ValueAdjusted As Double  'adjusted for surface area
  Public ARI As Double            'average return interval if > 1 year
  Public Duration As Double       'duration in timesteps
  Public EventSum As Double       'total precipitation volume of the event
End Class
