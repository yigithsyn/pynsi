from .Measurement import Measurement

import os

class BeamTable(Measurement):
  r""" Beam (Frequency) table operations

  Parameters
  ----------
  meas : Measurement
         selected measurement

  """
  def __init__(self, meas: Measurement):
    self.meas = meas

  def Count(self):
    r""" Returns the number of beams in the beam table.

    Returns
    ----------
    int
      number of beams (frequencies) in the beam table

    """
    self.meas.console.RunScriptFile("C:\\NSI2000\\Script\\ExportBeamCount.bas")
    beamCount = 0
    with open(self.meas.file.replace(".NSI","_BeamCount.txt")) as file:
      beamCount = int(file.readlines()[0])
    os.remove(self.meas.file.replace(".NSI","_BeamCount.txt"))
    return beamCount


  def SELECT_BEAM(self, beam: int) -> None:
    r""" Select beam from beam table

    Parameters
    ----------
    beam : int
          beam number
    """
    self.meas.console.SELECT_BEAM(beam) 

