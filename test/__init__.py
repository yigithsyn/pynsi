import sys, time
import pytest

sys.path.append('./src')
from pynsi import Measurement, FarField, Script, BeamTable, OEWGProbeType

# measurement = Measurement(r'C:\NSI2000\Data\pla11.nsi')
# farField    = FarField(measurement)
# script      = Script(measurement)
# beamTable   = BeamTable(measurement)

print("")
# print(measurement.file)
# script.runFile(r'C:\NSI2000\Script\HelloWorld.bas')
# script.importNFData(r"D:\ATAM\03 OLCUMLER\03 ANTEN\2024\057 ASELSAN\02 OLCUM DOSYALARI\02 ISIMA DESENI\SGH\NFScan.prm")

measurement = Measurement(r"D:\ATAM\03 OLCUMLER\03 ANTEN\2024\069 RST ELEKTRONIK\02 OLCUM DOSYALARI\02 ISIMA DESENI\66\NFScan.NSI")
beamTable   = BeamTable(measurement)
farField    = FarField(measurement)

print("- NSI2000 Professional Version: %s"%measurement.nsi2000Version())

farField.AZ_OVER_EL()
farField.L2_AZ_OVER_EL()
farField.HPTS(481)
farField.VPTS(481)
farField.PC_OEWG()
farField.PROBE_OEWG_TYPE(OEWGProbeType.WR90)

print("Number of beams: %d"%beamTable.Count())
for i in range(beamTable.Count()):
  print("Selecting beam number #%d"%(i+1))
  beamTable.SELECT_BEAM(i+1)
  print("- Frequency: %.1f MHz"%(measurement.MEAS_FREQUENCY()/1E6))
  print("- Farfield processing ... ", end="", flush=True)
  sys.stdout.flush()
  farField.NO_PLOT()
  print("Done")
  print("- Peak value: %.3f"%farField.PEAK())
  print("- Directivity: %.3f"%farField.DIRECTIVITY())
  print("- Gain: %.3f"%farField.CALCULATED_COMPARISON_GAIN())
  print("- Efficiency: %.3f"%farField.efficiency())
  print("- Peak position (H-cut): %.3f"%farField.PEAK_H())
  print("- Peak position (V-cut): %.3f"%farField.PEAK_V())

  farField.AUTO_H_CENTER_OFF()
  farField.AUTO_V_CENTER_ON()
  farField.HCUT()
  print("- Beamwidth (H-cut): %.3f"%farField.beamwidth())

  farField.AUTO_H_CENTER_OFF()
  farField.AUTO_V_CENTER_OFF()
  farField.VCUT()
  print("- Beamwidth (V-cut): %.3f"%farField.beamwidth())

  farField.RESET_CENTER()
  farField.AUTO_H_CENTER_OFF()
  farField.AUTO_V_CENTER_ON()
  farField.HCUT()
  sidelobes = farField.sidelobes()
  print("- Sidelobes (H-cut): %.3f, %.3f, %.3f, %.3f"%(sidelobes[0], sidelobes[1], sidelobes[2], sidelobes[3]))

  farField.RESET_CENTER()
  farField.AUTO_H_CENTER_OFF()
  farField.AUTO_V_CENTER_OFF()
  farField.VCUT()
  sidelobes = farField.sidelobes()
  print("- Sidelobes (V-cut): %.3f, %.3f, %.3f, %.3f"%(sidelobes[0], sidelobes[1], sidelobes[2], sidelobes[3]))

  # print("- Exporting pattern data ... ", end="", flush=True)
  # sys.stdout.flush()
  # farField.LISTING_TO_FILE()
  # print("Done")
  break
  time.sleep(0.1)

# farfield = measurement.FarField()

# assert print(pynsi.measurementFile()) == None

# assert pynsi.selectBeam(5) == 0


# # assert pynsi.importNFData("D:\\ATAM\\03 OLCUMLER\\03 ANTEN\\2024\\067 ERA HABERLESME\\02 OLCUM DOSYALARI\\02 ISIMA DESENI\\Downlink_LHCP\\NFScan.prm", 267, 267) == None
# # assert pynsi.runScriptFile("C:\\NSI2000\\Script\\ImportNFDataCurrent.bas") == None

# assert print(pynsi.measurementFile(r'C:\NSI2000\Data\pla11.nsi')) == None
