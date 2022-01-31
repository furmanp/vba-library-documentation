# Femap/Geometry

### *femap_xls_draw_points*
|femap_xls_draw_points <br> ()|
|:----------------------------------------|
|Description:|
|Method draws points from defined SET.ID's in a currently open FEMAP session. User is prompted with InputBox and asked to define set.id's he wants to have drawn in FEMAP. Appriopriate Excel data table is required for proper execution (See Remarks/Usage).|
|Input:|
|None|
|Output:|
|None|
|Return Code:|
|None|
|Remarks/Usage:|
|Method is primarily looking for **`set.id`** keyword, from which it will build 2D data array. <br> **`set.nam`** is obsolete for Femap and serves only to indicate the user what the points are referring to. **`pnt.id`** parameter defines unique ID that points will have assigned in Femap. If left empty, first available ID will be assigned to the points.<br> **`x, y, z`** stand for coordinates, provided in units defined in Femap session. <br> ![femap_xls_draw_points.png](../../assets/pictures/femap_xls_draw_points.png) <br> Method calls in **`femap_create_points, xls_inputbox, a_base0, a_make, a_tidy`**, and **`a_look`** functions.|
|Example:|
|None|

### *femap_create_points*
|femap_create_points <br> (x, y, z, pointID)|
|:----------------------------------------|
|Description:|
|Method creates points in currently active Femap session.|
|Input:|
|**DOUBLE `x, y, z`** &nbsp; Coordinates of the point to be created in Femap. Units should match as defined in Femap.|
|**LONG `[pointID]`** &ensp; Unique ID number of the points to be created. If omitted, or 0, femap will assign first available ID number to the point.|
|Output:|
|None|
|Return Code:|
|None|
|Remarks/Usage:|
|Function requires active Femap session.<br> Point ID's if defined has to be unique.|
|Example:|
|None|

### *femap_create_curve*

### *femap_draw_curve*