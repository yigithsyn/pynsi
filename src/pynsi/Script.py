from .Measurement import Measurement

import os


class Script(Measurement):
  r""" Script runner for general automation and file related processes

  Parameters
  ----------
  meas : Measurement
         selected measurement

  """
  def __init__(self, meas: Measurement):
    self.meas = meas

  def RunScriptFile(self, filename: str):
    r""" Run script file 

    Parameters
    ----------
    filename : str
               script file name with full path

    """
    if os.path.exists(filename):
      try:
        self.meas.console.RunScriptFile(filename)
      except:
        raise Exception
    else:
      raise FileNotFoundError
    