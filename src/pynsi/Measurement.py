import os
import win32com.client as win32
import pathlib


class Measurement:
  r""" Open measurement file or bind active measurement

  Parameters
  ----------
  filename : str
             measurement file path

  """
  def __init__(self, filename: str = None):
    self.server  = win32.gencache.EnsureDispatch('NSI2000.server')
    self.app     = self.server.AppConnection
    self.console = self.app.ScriptCommands

    if filename:
      if os.path.exists(filename):
        if pathlib.Path(filename).suffix == ".NSI" or pathlib.Path(filename).suffix == ".nsi":
          try:
            self.console.DATA_FILENAME = filename
            self.file = filename 
          except:
            raise Exception
        else:
          raise ValueError("Filename must be valid .NSI type")
      else:
        raise FileNotFoundError
    else:
      self.file = self.console.DATA_FILENAME

  def MEAS_FREQUENCY(self) -> float:
    r""" Returns beam frequency

    Returns
    -------
    float
      beam freqeuncy in Hertz [Hz]

    """
    return self.console.MEAS_FREQUENCY*1E9

  def CLEAR_ALL_PLOTS(self) -> None:
    r""" Clear all plot windows
    """
    return self.console.CLEAR_ALL_PLOTS()

  def nsi2000Version(self) -> str:
    r""" Returns NSI2000 software version

    Returns
    -------
    str
      version string
    """
    return self.app.Version



