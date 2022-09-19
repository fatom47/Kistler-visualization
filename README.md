# Kistler-visualization
Excel VBA visualization of Kistler pressing processes

Kistler pressing units generate their outputs in a specific format that makes them not really user friendly to process.  
The file contains information such as cluster and station and channel name, program, date and time, evaluation method, result, trigger, sensors, settings, offsets, reserved fields, evaluation windows, number of measurements, units, empty fields etc.  
After this ~150-line introduction, the measured values finally come.

Processing such files manually is time-consuming and therefore it directly encourages automatic processing.  
Let the mentioned macros be an example.

Instructions for use  
Display of pressing curves using macro Excel (.xlsm)

0) Close all Excel windows
1) Copy pressing processes files into the ./CSV/ subfolder
2) Run the .xlsm file and let Excel load the files  
-> if macros are not enabled, it is necessary to enable them  
-> evaluation window, X and Y axes, graph name are loaded form the first file, threfore do not mix different types of pressing (Gasoline/Diesel)  
-> the maximum number of curves that can be displayed at once is 255 (Excel restriction)
3) When finished, close Excel without saving the changes

To display only NOK curves, the deleteOKs.vba macro is used, which deletes the OK ones.
