This demo takes advantage of a weakness in VB; it doesn't notice if you have Private routines with the same name in different modules, so both modules contain the functions LightEffect and SafeRGBValue. 

This is a potential source of bugs in code, if the 2 routines are not exact copies of each other. 

In this demo they are exactly the same, I could also have renamed them by adding 'VB' or 'API' to end of names as i did with other routines (which are not identical but optimised (I think) for VB or API). 

Alternatively a single copy of the routines declared Public (in a 3rd module or either of the exiting modules) would work but would make the modules less portable.