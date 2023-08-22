# C#
Program currently works on certain conditions on Kyle's local machine.  Conditions:
1. Job is only for M-Units
2. Job us fir 0-10v DV Parameters
3. VFD Parameter sheets are in local drive at specified location
   
Project Next Step/Difficulty:
1.  De-Localize process, once it is safe to modify files beyond reading on CAS Drives - 1/10
2.  Make program usable with any job stlye - multi line/jobs with S-units - 4/10
3.  Make Program recognize VFD Parrameter style - EFI, RPC, 4-20mA, Potentiometer, Rotary Switch 3/10
4.  Fill in correct sheet with needed information for different VFD PArameter Styles - 7/10
5.  De-Localize any variables that are a constant with every Parameter - 5/10
6.  Add Error catching for human entry, including and not limited to: - 10/10
    characters enter into field
    files not found
    information not found on pdf
