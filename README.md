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

### Matrix
The library contains several different matrix factory methods:

 * ScalarMatrix(Value, m, n) - Create a matrix with m rows and n columns, with the value of Value in each element
 * Identity(m) - Create an identity matrix with M rows and columns
 * DiagonalMatrix(Vector) - Create a diagonal matrix with the elements of the supplied vector along the diagonal
 * MatrixFromJaggedArray(Array) - Create a matrix from a nested array of arrays.
 * CreateMatrix(Variant) - Create a matrix from a two-dimensional variant

Methods to create a new matrix:
```vb
Dim M as Matrix
' [0 0 0]
' [0 0 0]
Set M = ScalarMatrix(0, 2, 3)
```
```vba
' [1 0 0]
' [0 1 0]
' [0 0 1]
Set M = Identity(3)
```
```vba
' [2 3 3]
' [3 2 3]
' [3 3 2]
Set M = MatrixFromJaggedArray( _
            Array( _
                Array(2, 3, 3), _
                Array(3, 2, 3), _
                Array(3, 3, 2) _
            ) _
        )

Dim V as Vector
V = M.GetColumn(1)
' [2 0 0]
' [0 3 0]
' [0 0 3]
Set M = DiagonalMatrix(V)
```
```vba
Dim Mat As Variant
ReDim Mat(1 To 2, 1 To 2)
Mat(1, 1) = 1
Mat(1, 2) = 0
Mat(2, 1) = 0
Mat(2, 2) = 1
Dim M As Matrix
Set M = CreateMatrix(Mat)
```
The matrix class contains the following methods:
* isDiagonal()
* IsEqual(Matrix)
* Add(Matrix)
* Subtract(Matrix)
* Multiply(Matrix)
* ScalarMultiply(number)
* ScalarDivide(number)
* Transpose()
* GetRow(Integer)
* GetColumn(Integer)
* AugmentRight(Matrix)
* AugmentBelow(Matrix)
* ReplaceRow(Integer, Vector)
* ExcludeRow(Integer)
* Trace()
* GetDiagonalElements()
* SwapColumns(Integer, Integer)
* Submatrix(Integer, Integer, Integer, Integer)
* Inverse()
* ToJaggedArray()
* ToString()

### Vector
The library contains one Vector factory method:
 * CreateVector(Variant) - Create a vector from an array
```vba
Data = Array(1,2,3,4,5)

Dim V As Vector
Set V = CreateVertor(Data)
```
