from .Measurement import Measurement

import os

from enum import IntEnum

class ProbeType(IntEnum):
  NONE         = 0
  COSINE       = 1
  OEWG         = 2
  PATTERN_FILE = 3

class OEWGProbeType(IntEnum):
  WR1500 = 0
  WR975  = 1
  WR770  = 2
  WR650  = 3
  WR510  = 4
  WR430  = 5
  WR340  = 6
  WR284  = 7
  WR229  = 8
  WR187  = 9
  WR159  = 10
  WR137  = 11
  WR112  = 12
  WR90   = 13
  WR75   = 14
  WR62   = 15
  WR51   = 16
  WR42   = 17
  WR34   = 18
  WR28   = 19
  WR22   = 20
  WR19   = 21
  WR15   = 22
  WR12   = 23

class DisplayFormat(IntEnum):
  AMPLITUDE   = 0
  AXIAL_RATIO = 3

class FarField(Measurement):
  r""" Far field processing for selected measurement

  Parameters
  ----------
  meas : Measurement
         selected measurement

  """
  def __init__(self, meas: Measurement):
    self.meas = meas

  # Displayed polarization
  def EPRINC(self) -> None:
    r""" Sets the display pol flag to Co-pol (Pol-1)
    """
    self.meas.console.FF_EPRINC()

  def ECROSS(self) -> None:
    r""" Sets the display pol flag to Cross-pol (Pol-2)
    """
    self.meas.console.FF_ECROSS()


  # Output resolution
  def HPTS(self, number: int = None) -> int:
    r""" Sets/Gets number of HCut points

    Parameters
    ----------
    number : int
      number of resolution points

    Returns
    ----------
    int
      number of resolution points

    """
    if number: self.meas.console.FF_HPTS = number
    return self.meas.console.FF_HPTS
  
  def VPTS(self, number: int = None) -> int:
    r""" Sets/Gets number of VCut points

    Parameters
    ----------
    number : int
      number of resolution points

    Returns
    ----------
    int
      number of resolution points

    """
    if number: self.meas.console.FF_VPTS = number
    return self.meas.console.FF_VPTS

  # Plot center
  def RESET_CENTER(self):
    r""" Forces the H- and V-centers = 0.0
    """
    self.meas.console.FF_RESET_CENTER()
  def AUTO_H_CENTER_ON(self):
    r""" Forces H-peak position into the Hcenter
    """
    self.meas.console.FF_AUTO_H_CENTER_ON()
  def AUTO_V_CENTER_ON(self):
    r""" Forces V-peak position into the Vcenter
    """
    self.meas.console.FF_AUTO_V_CENTER_ON()
  def AUTO_H_CENTER_OFF(self):
    r""" Forces H-peak position = 0
    """
    self.meas.console.FF_AUTO_H_CENTER_OFF()
  def AUTO_V_CENTER_OFF(self):
    r""" Forces V-peak position = 0
    """
    self.meas.console.FF_AUTO_V_CENTER_OFF()

  # Coordinate system select routines
  def AZ_OVER_EL(self):
    r""" Sets the coordinate system to Az/El
    """
    self.meas.console.FF_AZ_OVER_EL()
  def EL_OVER_AZ(self):
    r""" Sets the coordinate system to El/Az
    """
    self.meas.console.FF_EL_OVER_AZ()
  def THPH(self):
    r""" Sets the coordinate system to Theta/Phi
    """
    self.meas.console.FF_THPH()

  # Farfield sense select routines
  def TAU(self, tau: float = None):
    r""" returns or sets the far-field Tau angle
    """
    if tau != None: self.meas.console.FF_TAU = tau
    else: return self.meas.console.FF_TAU
  def AUTO_TAU_AT_PEAK(self, state: bool = True):
    r""" sets a flag to automatically calculate the Tau angle
    """
    self.meas.console.FF_AUTO_TAU_AT_PEAK = state
  def REFERENCE_POL_MATCHES_MEASPOL1(self, state: bool = True):
    r""" sets a flag to force the Tau angle to match the measurement Pol-1 sense
    """
    self.meas.console.FF_REFERENCE_POL_MATCHES_MEASPOL1 = state
  def LINEAR_POL(self):
    r""" sets the far-field sense flag = Linear
    """
    self.meas.console.FF_REFERENCE_POL_MATCHES_MEASPOL1 = True
    self.meas.console.FF_LINEAR_POL()
  def RHCP(self):
    r""" sets the far-field sense flag = RHCP
    """
    self.meas.console.FF_REFERENCE_POL_MATCHES_MEASPOL1 = False
    self.meas.console.FF_RHCP()
  def LHCP(self):
    r""" sets the far-field sense flag = LHCP
    """
    self.meas.console.FF_REFERENCE_POL_MATCHES_MEASPOL1 = False
    self.meas.console.FF_LHCP()

  # Polarization select routines
  def L2_AZ_OVER_EL(self):
    r""" Sets the polarization flag to EazEel
    """
    self.meas.console.FF_L2_AZ_OVER_EL()
  def L2_EL_OVER_AZ(self):
    r""" Sets the polarization flag to EalEep
    """
    self.meas.console.FF_L2_EL_OVER_AZ()
  def L2_ETH_EPH(self):
    r""" Sets the polarization flag to EalEep
    """
    self.meas.console.FF_L2_ETH_EPH()
  def LUDWIG3(self):
    r""" Sets the coordinate system to L3
    """
    self.meas.console.FF_LUDWIG3()

  # Probe select routines
  def PC_OEWG(self) -> None:
    r""" Sets the probe correction model to Open ended waveguide (OEWG)
    """
    self.meas.console.FF_PC_OEWG()
  def PC_NONE(self) -> None:
    r""" Sets the probe correction model to None (No probe correction)
    """
    self.meas.console.FF_PC_NONE()
  def PC_COSINE(self) -> None:
    r""" Sets the probe correction model to Cosine pattern
    """
    self.meas.console.FF_PC_COSINE()
  
  def PROBE_OEWG_TYPE(self, oewg: OEWGProbeType = None):
    r""" Returns or sets the OEWG probe model

    Returns
    -------
    int
      OEWG probe model

      WR1500 = 0
      WR975  = 1
      WR770  = 2
      WR650  = 3
      WR510  = 4
      WR430  = 5
      WR340  = 6
      WR284  = 7
      WR229  = 8
      WR187  = 9
      WR159  = 10
      WR137  = 11
      WR112  = 12
      WR90   = 13
      WR75   = 14
      WR62   = 15
      WR51   = 16
      WR42   = 17
      WR34   = 18
      WR28   = 19
      WR22   = 20
      WR19   = 21
      WR15   = 22
      WR12   = 23
    """
    self.meas.console.FF_PC_OEWG()
    if oewg: self.meas.console.FF_PROBE_OEWG_TYPE = oewg
    return self.meas.console.FF_PROBE_OEWG_TYPE

  def DISPLAY_SOURCE(self, format: DisplayFormat = None) -> DisplayFormat:
    r""" The options control which source (amp or phase) will be plotted when making a plot
    """
    if format: self.meas.console.FF_DISPLAY_SOURCE = format
    return self.meas.console.FF_DISPLAY_SOURCE
    
  
  def NO_PLOT(self):
    r""" Produce far-field data without making a plot
    """
    self.meas.console.FF_NO_PLOT()


  def PEAK(self, fast_process: bool = False) -> float:
    r""" Gets the “Global Peak” which is the peak value of both pols no matter what the normalization is set to
    
    Parameters
    ----------
    fast_process : bool
      far farfield processing for global peak search default = True

    Returns
    -------
    float
      global peak of farfield
    """
    if fast_process:
      hPts = self.HPTS()
      vPts = self.VPTS()
      self.HPTS(2)
      self.VPTS(2)
      self.meas.console.FF_NO_PLOT()
    peak = self.meas.console.FF_PEAK
    if fast_process:
      self.HPTS(hPts)
      self.VPTS(vPts)
    return peak
  
  # Farfield cuts
  def HCUT(self):
    r""" Makes a far-field H-cut
    """
    self.meas.console.FF_HCUT()

  def VCUT(self):
    r""" Makes a far-field V-cut
    """
    self.meas.console.FF_VCUT()

  # Beam parameters extract routines
  def PEAK_H(self) -> None:
    r""" Returns H-axis FF peak position

    Returns
    -------
    float
      H-axis FF peak position
    """ 
    return self.meas.console.FF_PEAK_H   
  
  def PEAK_V(self) -> None:
    r""" Returns V-axis FF peak position

    Returns
    -------
    float
      V-axis FF peak position
    """ 
    return self.meas.console.FF_PEAK_V  
  
  def beamwidth(self, beamwidthLevel: float = -3) -> float:
    r""" Returns beamwidth of last displayed cut 

    Parameters
    ----------
    beamwidthLevel : float
      beamwidth level to be calculated

    Returns
    -------
    float
      beamwidth
    """ 
    self.meas.console.FF_BEAMWIDTH_DB1 = beamwidthLevel
    return self.meas.console.FF_BEAMWIDTH_VAL1  
  
  def sidelobes(self) -> tuple:
    r""" Returns sidelobes of last displayed cut

    Returns
    -------
    tuple
      left position, left value, right position, right value
    """ 
    return (self.meas.console.FF_SIDELOBE_LEFT_LOC, self.meas.console.FF_SIDELOBE_LEFT_VAL, self.meas.console.FF_SIDELOBE_RIGHT_LOC, self.meas.console.FF_SIDELOBE_RIGHT_VAL)  
  
  # Gain calibration routines
  def UPDATE_CAL_FROM_CAL_FILE(self, state: bool = True) -> None:
    r""" Sets the flag to update the SGA max far-field (FF_SGA_PEAK) from the file
    
    Parameters
    ----------
    state : bool
      update gain from calibration file
    """
    self.meas.console.FF_UPDATE_CAL_FROM_CAL_FILE = state
  def CAL_FILENAME(self, filename: str) -> None:
    r""" Sets the calibration filename that stores the gain reference max far-field values
    
    Parameters
    ----------
    filename : bool
      gain reference peak values stored in .ncl file wihch is exported by ExportGlobalPeak.bas macro
    """
    if os.path.exists(filename):
      self.meas.console.FF_CAL_FILENAME             = filename
      self.meas.console.FF_UPDATE_CAL_FROM_CAL_FILE = True
    else:
      raise FileNotFoundError

  def UPDATE_GAIN_FROM_GAIN_TABLE_FILE(self, state: bool = True) -> None:
    r""" Sets the flag to update the SGA max far-field (FF_SGA_PEAK) from the file
    
    Parameters
    ----------
    state : bool
      update gain from gain table file
    """
    self.meas.console.FF_UPDATE_GAIN_FROM_GAIN_TABLE_FILE = state
  def GAIN_TABLE_FILENAME(self, filename: str) -> None:
    r""" Sets gain table filename that stores the obtained gain values of reference
    
    Parameters
    ----------
    filename : bool
      obtained gain values stored in .ngt
    """
    if os.path.exists(filename):
      self.meas.console.FF_GAIN_TABLE_FILENAME              = filename
      self.meas.console.FF_UPDATE_GAIN_FROM_GAIN_TABLE_FILE = True
    else:
      raise FileNotFoundError
    
  def DIRECTIVITY(self) -> None:
    r""" Returns directivity

    Returns
    -------
    float
      directivity in decibels [dBi]
    """ 
    return self.meas.console.FF_DIRECTIVITY 
    
  def CALCULATED_COMPARISON_GAIN(self) -> None:
    r""" Returns calculated gain comparison value

    Returns
    -------
    float
      gain in decibels [dBi]
    """ 
    return self.meas.console.FF_CALCULATED_COMPARISON_GAIN
    
  def efficiency(self) -> None:
    r""" Returns total efficiency

    Returns
    -------
    float
      efficiency in descibels [dB]
    """ 
    return self.meas.console.FF_CALCULATED_COMPARISON_GAIN - self.meas.console.FF_DIRECTIVITY 

  
  def LISTING_TO_FILE(self):
    r""" Writes the far-field data to the specified file for both polarizations 
    """
    filename = self.meas.file.replace(".NSI","_Beam%d.txt"%(self.meas.console.BEAM_NUMBER+1))
    self.meas.console.FF_LISTING_TO_FILE(filename, True)



