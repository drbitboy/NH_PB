"""
Expect to see in LSK (*.tls):

  DELTET/DELTA_T_A       =   32.184
  DELTET/K               =    1.657D-3
  DELTET/EB              =    1.671D-2
  DELTET/M               = (  6.239996D0   1.99096871D-7 )
  DELTET/DELTA_AT        = ( 10,   @1972-JAN-1 ... )

Expect to see in SCLK (*.tsc):

  SCLK_DATA_TYPE_98          = (        1 )
  SCLK01_TIME_SYSTEM_98      = (        2 )
  SCLK01_N_FIELDS_98         = (            2 )
  SCLK01_MODULI_98           = (   4294967296       50000 )
  SCLK01_OFFSETS_98          = (            0           0 )
  SCLK01_OUTPUT_DELIM_98     = (            2 )
  SCLK_PARTITION_START_98    = ( 0.00000000000000e+00 ... )
  SCLK_PARTITION_END_98      = ( 7.01906785000000e+12 ... )
  SCLK01_COEFFICIENTS_98     = ( ... )

Typical usage:

  python met2utc.py [naif0013.tls] [new-horizons_3382.tsc] > met2utc.csv

"""
import sys
import numpy
import datetime as dt
import spiceypy as sp

if "__main__" == __name__:
  try:
    ### Default kernels
    kernels = dict()
    for kfn in 'naif0012.tls new-horizons_3381.tsc'.split() + sys.argv[1:]:
      kernels[kfn.split('.')[-1]] = kfn

    ### Column headers ...
    row1 = 'met,tdt-leapsec,offset_scale,utc,tdt'
    ### ... plus other kernels from command line
    for key in kernels:
      kfn = kernels[key]
      row1 = row1 + f',"{kfn}"'
      sp.furnsh(kfn)

    ### Get parameters from LSK:  32.184 and last leapsecond pair
    delta_t_a = sp.gdpool('DELTET/DELTA_T_A',0,1)[0]
    delta_at,last_leapsec_tdt = sp.gdpool('DELTET/DELTA_AT',0,999)[-2:]

    ### Get parameters from SCLK
    ### - Piecewise-linear coefficients:  [tiks,TDT,rate]*N
    coeffs = numpy.array(sp.gdpool('SCLK01_COEFFICIENTS_98',0,20000)).reshape((-1,3,))

    iw = numpy.where(coeffs[:,1] > last_leapsec_tdt)
    coeffs = coeffs[iw]

    ### - Modulus to convert from tiks to MET
    moduli = sp.gdpool('SCLK01_MODULI_98',1,99)
    moduli_1 = 1.0
    for modulus in moduli: moduli_1 *= modulus
    coeffs[:,0] /= moduli_1

    ### Add equiv. of 20y of MET to last coeffs row (tiks/modulus,tdt)
    ### to model (extrapolate) times after that last row
    twentyy = 20. * 365.25 * sp.spd()
    newrow = numpy.array([coeffs[-1,0] + twentyy
                         ,coeffs[-1,1]+(twentyy * coeffs[-1,2])
                         ,coeffs[-1,2]
                         ]).reshape((-1,3,))
    coeffs = numpy.vstack((coeffs,newrow,))

    ### Row 1:  Column headers, plus kernels
    print(row1)

    ### Calculate row number of last row of coeffs
    lastrow = coeffs.shape[0] + 2

    ### Row 2:  coeffs cell ranges (A, B), offset and scale (C)
    print(f'"A3:A{lastrow}","B3:B{lastrow}","C3:C4"')

    ### scale seconds/day, Time Offset => cells C4 and C3, respectively
    cpops = [sp.spd(),'"2000-01-01 12:00:00"']

    ### Calculate leapsecond offset TDT => UDT
    delta = delta_t_a + delta_at

    ### Write coefficients i.e. (MET,TDT) pairs into Columns A and B
    for met,tdt,rate in coeffs:

      ### Column C data (offset and scale)
      try   : cpop = cpops.pop()
      except: cpop = ''

      ### Adjust leapsecond offset to hack  equivalent UTC
      tdt -= delta

      ### Get UTC and TDT strings, for checking
      tdb = sp.sct2e(-98,met*5e4)
      utc = sp.timout(tdb,'"YYYY-MM-DDTHR:MN:SC.#######Z" ::UTC')
      stdt = sp.timout(tdb,'"YYYY-MM-DD-HR:MN:SC.#######" ::TDT')

      ### print the row as a CSV
      print(f"{met:.5f},{tdt:.6f},{cpop},{utc},{stdt}")

  except:
    import traceback
    traceback.print_exc()
    print(__doc__)
