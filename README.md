VBA Math Objects
=====================

### An advanced object oriented math library for VBA

Features
--------
 * [Matrix](#matrix)
 * [Vector](#vector)
 * [Unit Tests](#unit-tests)

Upcoming
--------
 * [Regression](#regression)
 * [Smoothing](#smoothing)
 * [Interpolation](#interpolation)
 
 Setup
-----

Open Microsoft Visual Basic For Applications and import each cls and bas and frm file into a new project. Name the project VBA-Math-Objects and save it as an xlam file. Enable the addin. Within Microsoft Visual Basic For Applications, select Tools>References and ensure that  "Microsoft ActiveX Data Objects x.x Library", "Microsoft Scripting Runtime", and VBA-Math-Objects is selected.

 Testing
 -----
The unit tests demonstrate many ways to use each of the classes. To run the tests, Import all the modules from the testing directory into a spreadsheet, install the [VBA-Unit-Testing library](https://github.com/Beakerboy/VBA-Unit-Tester) and type '=RunTests()' in cell A1. Ensure the Setup steps have all been successfully completed for that library.
 
 Usage
-----
