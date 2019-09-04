
<h3> About </h3>

<p align="justify">
ShapeFactory library is intended to facilitate the usage of high-performance numerical methods in architecture and design.
Based on IDispatch interface it removes all unnecessary features, such as support for multiple platforms, or creation of COM servers.
Roughly speaking this library address specific CATIA platform, instead of trying to incorporate entire COM implementation.
Compared to more general win32com libraries, this means better support for CATIA automation developers and ability to do breaking changes
for the sake of better CATIA programming experience. Also, due to Julia transparent interaction with C or Fortran, all enterprise knowledge
can be represented in clean and fast code, without annoying boilerplate or heavy and slow frameworks. Generative design field deals mainly
with computational methods, this is somewhat different from software development, therefore it is better to use specialized platform full of
sophisticated algorithms. </p>

<h3> Installation </h3>

```
pkg> add ShapeFactory
```

<h3> Examples </h3>

System initialization

```julia
using ShapeFactory
cat = dispatch("CATIA.Application")
fac = cat.ActiveDocument.Part.HybridShapeFactory
cat.StatusBar = "simple example"
```

Random points in xy plane

```julia
r = () -> rand(1:100, 20)
fac.AddNewPointCoord.(r(), r(), 0)
```

Concentric circles

```julia
pln = prt.FindObjectByName("xy plane")
cen = fac.AddNewPointCoord(0, 0, 0)
con = fac.AddNewCircleCtrRad.(cen, pln, false, 10:10:100)
```

As you can see, all constructors vectorized and entire library preserve object oriented paradigm of CATIA Automation API.

<h3> More examples </h3>

More examples can be found in \examples directory of the master branch.
