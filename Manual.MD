---
#### cPCA.cls
---
This class is used to run principal component analysis.

Requires: [modMath.bas](Modules/modMath.bas)

##### Available methods:

```Ruby
PCA(x() As Double, Optional first_n As Long = 0, Optional use_SVD As Boolean = False, Optional save_rawdata As Boolean = False)
```

**Desc:** Perform PCA transformation on data x()

**Input:** 
- x(),          a NxD array of data, where N is number of samples and D is dimension of data  
- first_n,      number of PCs to calculate. If left at zero, all componets are calculated.  
                      Otherwise only first_n PCs are calculated using power iteration.  
- use_SVD,      if set to TRUE, singular value decompositon is performed on x(). Otherise  
                      covariance matrix of x() is calculated and Jacobi method is used.  
- save_rawdata, if set to TRUE then x() will be saved in class object and can be accessed by method .raw_data  

```Ruby
Reset()
```

**Desc:**   Release memory in class object

```Ruby
Vec(Optional n_vec as variant) As Double()
```

**Desc:**   Read only property, return the first n_vec unit-vectors of PCs, in an array of size Vec(1:D, 1:n_vec)

**Input:**  n_vec, number of vectors to return. If left blank then all calculated eigenvectors are returned

```Ruby
Val(Optional n_vec as variant) As Double()
```

**Desc:**   Read only property, return the first n_vec eigenvalues, in an array of size Val(1:n_vec)

**Input:**  n_vec, number of values to return. If left blank then all calculated eigenvalues are returned

```Ruby
x_PCA(Optional n_dim As Variant) As Double()
```

**Desc:**  Read only property, return projections of raw data on the first n_dim PCs, in an array of size(1:N, 1:n_dim)

**Input:** n_dim, number of components to return. If left blank then all calculated PCs are returned

```Ruby
BiPlot_Print(vRng As Range, Optional PC1 As Long = 1, Optional PC2 As Long = 2, Optional magnify As Double = 1)
```

**Desc:**   output data necessary to chart a biplot in Excel. First two column are scatter chart,
        next two columns are chart with scatter plot with straight lines.
        
**Input:**
 - vRng,       upper-left cell of output range
 - PC1 & PC2,  components to print on the x & y axis respectively, default are the 1st and 2nd components.
 - magnify,    default value is 1 and a vector is shown at length equals to square root of its eigenvalue.
                    The length can be mutiplied by magnify so it can be chart together with x_PCA().
