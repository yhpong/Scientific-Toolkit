SCreenshots\# SciTool
This is a VBA library of basic algorithms commonly used in data analysis. Althouhg there are many state of the art implementations of these algorithms in other languages like Matlab, R or Python, they are not often seen in VBA. While are good reasons to not use Excel or VBA for these type of analysis, but if you are stuck with Excel either becasue of budget issue or IT policy issue, then hopefully this library can offer a bit of help.
The librarary is still an ongoing project. Better documentations will come in time.
In this Readme, I will showcase some capabilities of what can be done with the library.

Test data here is wine data set from [UCI Machine Learning Datasets](https://archive.ics.uci.edu/ml/datasets.html). It consists of 178 samples of wines collected from three different cultivars, which will be named as W1, W2 and W3 in the remaning section. 13 attributes of these wine samples are measured.

1. Forina, M. et al. [UCI Machine Learning Repository](http://archive.ics.uci.edu/ml). Institute of Pharmaceutical and Food Analysis and Technologies. 

## Unsupervised Learning
Let's we are given these sample of wines, without knowing where they from. So we measure the 13 attributes of these samples, ranging from alchohol content to color intenisty. From the measurements we want to discover possible ways to classified these samples.

First we will import the data, the data should take the form of an array x() of size N X D, where N=178 is the number of samples, D=13 is the number of dimension.

Then the data first needs to be normalized. We will use zero mean and unit varaince in this case. The syntax will be:
```
Call modmath.Normalize_x(x,x_mean,x_sd,"AVGSD")
```


### Principal Component Analysis
Requires: cPCA.cls
![PCA](Screenshots/PCA.jpg)

```
Dim PCA1 as new cPCA
With PCA1
    Call .PCA(x)                            'Perform PCA transformation
    x_projection=.x_PCA(2)                  'output first two components
    Call .BiPlot_Print(Range("I3"), 1, 2)   'output biplot of component 1 & 2
End with
```
The method .PCA performs a transformation on x. The transformed data can then be extracted with method .x_PCA. In this case the first 2 components are saved to x_projection, which is shown in the left chart above. We also output the biplot of PC1 and PC2 to cell I3, which can be chart in Excel in a normal way, shown on the right hand side.

### t-SNE (t-Distributed Stochastic Neighbor Embedding)
Requires: ctSNE.cls, cqtree.cls, cqtree_point.cls, mkdtree.bas
![tSNE](Screenshots/tSNE.jpg)
```
Dim TS1 As New ctSNE
With TS1
    Call .tSNE_BarnesHut(x, 2)  'Perform t-SNE on raw data onto 2-dimension
    y = .Output                 'Output 2D projection of data
    z = .cost_function(True)    'output cost function to see convergence
    Call .Reset                 'release memory
End With
```
There are two methods in this class to perform transformation: .tSNE or .tSNE_BarnesHut. .tSNE is the simplest implementation of the algorithm. .tSNE_BarnesHut uses a quadtree data structure to speed up the process when number of data points N is huge. When N is small, the overhead cost of BarnesHut maynbot be worth the effort. But for large N~1000 , BarnesHut is essential for a resonable excecution time. The method .Output extract the transformed data which is plot in the above figures.
Note that random initialization is implemented, and different realizations will converge to different results, which also depends the hyperparameters used. The two figures above are two different run of t-SNE. Although the charts look different, they both produce similar relative ordering of data.
