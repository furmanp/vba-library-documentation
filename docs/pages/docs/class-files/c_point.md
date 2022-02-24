# cPoint 

cPoint class aims to represent objects that will simplify handling geometrical data. Structure of cPoint mimics the points structure
that is used in FEMAP. 

cPoint object has following properties:

* **`LONG ID`** - value that represents unique identification number of each point
* **`INTEGER X`** - value representing X coordinate of the point
* **`INTEGER Y`** - value representing Y coordinate of the point
* **`INTEGER Z`** - value representing Z coordinate of the point

Class is initialized with 0 values for all properties. 
To initialize the class objects with defined parameters follow example below:

```
Dim point As New cPoint

point.X = 5
point.Y = 7
point.Z = 14
point.ID = 1
```
