# C#
Program currently works on certain conditions on local machine.  Conditions:
1. Job is only for M-Units
2. Job is for 0-10v DC Parameters
3. VFD Parameter sheets are in local drive at specified location
   
Project Next Step/Difficulty:
1.  De-Localize process, once it is safe to modify files beyond reading on CAS Drives - 1/10
2.  Make program usable with any job stlye - multi line/jobs with S-units - 4/10
3.  Make Program recognize VFD Parrameter style - EFI, RPC, 4-20mA, Potentiometer, Rotary Switch 3/10
4.  Fill in correct sheet with needed information for different VFD Parameter Styles - 7/10
5.  De-Localize any variables that are a constant with every Parameter - 5/10
6.  Add Error catching for human entry, including and not limited to: 10/10 <br>
    characters enter into field<br>
    files not found<br>
    information not found on pdf<br>
    Possible converting null literal or possible null value to non-nullable type<br>
    failsafe workbook.Close() for program failure/crash<br>
<br>
Estimated Finish Time: 3 Weeks
