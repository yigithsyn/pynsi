from .Measurement import Measurement

class NearField(Measurement):
  r""" Near field processing for selected measurement

  Parameters
  ----------
  meas : Measurement
         selected measurement

  """
  def __init__(self, meas: Measurement):
    self.meas = meas

  def LISTING_TO_FILE(self):
    r""" Writes the far-field data to the specified file for both polarizations 
    """
    filename = self.meas.file.replace(".NSI","_Beam%03d_NF.txt"%(self.meas.console.BEAM_NUMBER+1))
    self.meas.console.NF_LISTING_TO_FILE(filename, True)
