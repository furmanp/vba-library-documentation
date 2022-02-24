# cCurve 

cCurve class aims to represent objects that will simplify handling geometrical data. Structure of cCurve mimics the feCurve structure
that is used in FEMAP. 

 
cCurve object has following properties:

* **`LONG ID`** - value that represents unique identification number of each curve
* **`Collection points`** - collection of `cPoint` objects which build the single curve.


Class is initialized with 0 value for ID and empty collection of points.
To initialize the class objects with defined parameters follow example below:
```
Dim point As New cPoint
Dim curve As New cCurve

point.X = 5
point.Y = 7
point.Z = 14
point.ID = 1
    
curve.points.Add point
curve.ID = 9
```

!!! note 
    There is no upper/lower limit in number of **`cPoint`** objects that are stored in single **`cCurve`** object. 
    Keep that in mind when forwarding the object further to FEM/CAD program as it may prompt an error 
    if **`cCurve`** stores only one point, which mathematically is not a curve. 
    Same relates to the upper limit, some programs have cap of 100 points per curve. 


### *ExportPoints*
|ExportPoints <br> this.ExportPoints|
|:----------------------------------------|
|Description:|
|Method converts collection of **`cPoints`** of cCurve object that is calling the function and returns a two dimensional array **`VARIANT/DOUBLE`** of coordinates|
|Input:|
|**cCurve `object`** &nbsp; cCurve objects only have access to this method, no other value is possible.|
|Output:|
|**VARIANT `object.ExportPoints`** function does not affect collection that is stored in the cCurve object. Return value has to be assigned to new `VARIANT` variable.|
|Return Code:|
|None|
|Remarks/Usage:|
|If **`cCurve`** object will be empty/will have no points stored, function will return value **-1**.|
|Example:|
|<pre><code>Dim curve As New cCurve<br>Dim point As New cPoint<br>Dim a_points As Variant<br><br>a_points = curve.ExportPoints<br></code></pre> |
