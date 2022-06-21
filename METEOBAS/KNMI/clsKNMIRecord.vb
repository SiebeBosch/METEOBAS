Public Class clsKNMIRecord
  Public DateInt As Integer
  Public HourInt As Integer
  Public NSLRAW As Integer    'precipitation value in 1/10 mm
  Public NSLCOR As Single     'corrected precipitation in mm
  Public MAKRAW As Integer    'makkink evaporation in 1/10 mm
  Public PMRAW As Integer     'penman monteith evaporation in 1/10 mm
  Public EventSum As Single   'total precipitation volume (mm) for the underlying event
  Public Duration As Integer  'duration of the underyling event (h)
  Public ARI As Integer       'average recurrance interval for the underlying event (y)

  Public alpha As Double      'parameter for the probability distribution
  Public kappa As Double      'parameter for the probability distribution
  Public mu As Double         'parameter for the probability distribution

End Class
