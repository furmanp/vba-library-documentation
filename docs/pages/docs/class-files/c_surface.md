# cSurface 

cSurface class aims to represent objects that will simplify handling geometrical data. Structure of cSurface mimics the feSurface structure
that is used in FEMAP. 

 
cSurface object has following properties:

* **`LONG ID`** - value that represents unique identification number of each curve
* **`Collection curves`** - collection of `cCurve` objects which build the single curve.


Class is initialized with 0 value for ID and empty collection of points.
To initialize the class objects with defined parameters follow example below:
```
Dim point As New cPoint
Dim curve As New cCurve
Dim surface as New cSurface

point.X = 5
point.Y = 7
point.Z = 14
point.ID = 1
    
curve.points.Add point
curve.ID = 9

surface.curves.Add curve
surface.ID = 1001
```

!!! note 
    There is no upper/lower limit in number of **`cCurve`** objects that are stored in single **`cSurface`** object. 
    Keep that in mind when forwarding the object further to FEM/CAD program as it may prompt an error 
    if **`cSurface`** stores only one curve, which mathematically is not a surface. 
    

### *ExportCurves*
|ExportCurves <br> this.ExportCurves|
|:----------------------------------------|
|Description:|
|Method converts collection of **`cCurves`** of **`cSurface`** object that is calling the function and returns a vector of 2 dimensional arrays **`VARIANT/DOUBLE`** that stores curves and points coordinates|
|Input:|
|**cSurface `object`** &nbsp; cSurface objects only have access to this method, no other value is possible.|
|Output:|
|**VARIANT `object.ExportCurves`** function does not affect collection that is stored in the cCurve object. Return value has to be assigned to new `VARIANT` variable.|
|Return Code:|
|None|
|Remarks/Usage:|
|If **`cSurface`** object will be empty/will have no points stored, function will return value **-1**.|
|Example:|
|<pre><code>Dim surface As New cSurface<br>Dim a_curves As Variant<br><br>a_curves = surface.ExportCurves<br></code></pre>|
