---
### cPCA.cls
---
This class is used to run principal component analysis.

Requires: [modMath.bas](Modules/modMath.bas)

#### Available methods:

```Ruby
PCA(x() As Double, Optional first_n As Long = 0, Optional use_SVD As Boolean = False)
```

**Desc:** Perform PCA transformation on data x()

**Input:** 
- x(),          a NxD array of data, where N is number of samples and D is dimension of data  
- first_n,      number of PCs to calculate. If left at zero, all componets are calculated.  
                      Otherwise only first_n PCs are calculated using power iteration.  
- use_SVD,      if set to TRUE, singular value decompositon is performed on x(). Otherise  
                      covariance matrix of x() is calculated and Jacobi method is used.  

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
x_PCA(Optional n_vec As Variant) As Double()
```

**Desc:**  Read only property, return projections of raw data on the first n_vec PCs, in an array of size(1:N, 1:n_vec)

**Input:** n_vec, number of components to return. If left blank then all calculated PCs are returned

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




---
### ctSNE.cls
---
This class is used to run t-SNE (t-Distributed Stochastic Neighbor Embedding). Main reference includes:
- "Visualizing High-Dimensional Data Using t-SNE", Laurens van der Maaten (2008)
- "Accelerating t-SNE using Tree-Based Algorithms", Laurens van der Maaten (2014)
- Implementations can be found on the author's [page](https://lvdmaaten.github.io/tsne/)

Requires: [mkdTree](Modules/mkdtree.bas), [cqtree](Modules/cqtree.cls), [cqtree_point](Modules/cqtree_point.cls)

#### Available methods:

```Ruby
tSNE(x() As Double, tgt_dimension As Long, _
            Optional perplexity As Double = 30, Optional perp_err As Double = 0.0001, _
            Optional learn_rate As Double = 100, Optional momentum As Double = 0.5, Optional max_iterate As Long = 1000, _
            Optional input_dist As Boolean = False)
```
**Desc:** Project raw data x() onto tgt_dimension.

**Input:** 
- x(),          a NxD array of data, where N is number of samples and D is dimension of data  
- tgt_dimension,  number of output deimsnion, usually it's 2.
- perplexity
- perp_err
- learn_rate
- momentum
- max_iterate
- input_dist, set to TRUE if x() is already a pairwsie distance matrix of size NxN

```Ruby
tSNE_BarnesHut(x() As Double, tgt_dimension As Long, _
            Optional perplexity As Double = 30, Optional perp_err As Double = 0.0001, _
            Optional learn_rate As Double = 100, Optional momentum As Double = 0.5, Optional max_iterate As Long = 1000, _
            Optional input_dist As Boolean = False)
```
**Desc:** Project raw data x() on tgt_dimension. Use tree to accelerate the process. Suitable for large number of samples.

**Input:**  Same as tSNE. But tgt_dimension only supports 2D at the moment.

```Ruby
Reset()
```
**Desc:** Release memory from class object

```Ruby
cost_function(Optional show_index As Boolean = False) As Double()
```
**Desc:** Read only property, return cost fucntion when tSNE was run

**Input:** show_index, if set to false, only return a vector. If set to false, return a 2D array with iterate number in first column.

```Ruby
Output() As Double()
```
**Desc:** Read only property, return porjections as an array of size (1:N, 1:tgt_dimension)


---
### ckMeanCluster.cls
---
This class is used to run k-Means clustering. [k++ scheme](https://en.wikipedia.org/wiki/K-means%2B%2B) is implemented for initialization.

#### Available methods:

```Ruby
kMean_Clustering(x() As Double, k As Long, Optional iterate_max As Long = 100, Optional strType As String = "EUCLIDEAN")
```
**Desc:** Perform k-Means clustering on raw data x().

**Input:** 
- x(),            a NxD array of data, where N is number of samples and D is dimension of data  
- k, target number of clusters
- iterate_max, maximum allowed number of iterations
- strType, type of distance metrics to use. Supports "EUCLIDEAN", "CORREL"

```Ruby
k_cluster() As Long
```
**Desc:** Read only property, returns number of clusters.

```Ruby
cluster_mean() As Double()
```
**Desc:** Read only property, returns mean of each cluster as an array of size (1:k,1:D)

```Ruby
cluster_size() As Long()
```
**Desc:** Read only property, returns size of each cluster as an integer vector of size (1:k)

```Ruby
x_cluster() As Long()
```
**Desc:** Read only property, returns clsuter index of each data point as an integer vector of size (1:N)

```Ruby
Reset()
```
**Desc:** Release memory from class


---
### cGraphAlgo.cls
---
This class is used to build undirected graph, using edge list as the main data structure.
For a graph with m edges, 
- pEdgeList(): a m x 2 array storing the start and end node index of each edge
- pEdgeDist(): a m length vector storing the length of each edge

#### Available methods:

```Ruby
MST_Build(x_dist() As Double)
```
**Desc:** Build Minimum Spanning Tree (MST) from a pairwise distance matrix x_dist().

**Input:** x_dist(),  NxN pairwise distance matrix for N nodes.

```Ruby
PMFG_Build(x_dist() As Double)
```
**Desc:** Build Planar Maximally Filterd Graph (PMFG) from a pairwise distance matrix x_dist().

**Input:** x_dist(),  NxN pairwise distance matrix for N nodes.


```Ruby
Init(EdgeList() As Long, EdgeDist() As Double, node_pos() As Double)
```
**Desc:** Build graph from a previously saved edge list, edge distance and node layout

**Input:** 
- EdgeList(), integer matrix of size m x 2, storing the ending points of m edges.
- EdgeDist(), real vector of size m, storing the length of each edge.
- node_pos(), real matrix of size n x 2, storing the 2D layout of n nodes.

```Ruby
Copy(g As cGraphAlgo)
```
**Desc:** Copy all attribules from anotehr graph 

**Input:** g, a cGraphAlgo object

```Ruby
Reset()
```
**Desc:** Release memory from class

Read-Only Properties:
```
Size() As Long                'number of nodes
n_edge() As Long              'number of edges
EdgeList() As Long()          'edge list, array of size (1:n_edge, 1:2)
EdgeDist() As Double()        'edge length, vector of length (1:n_edge)
node_degree() As Long()       'degree of each node, vector of length (1:Size)
node_degree_wgt() As Double() 'weighted degree of each node, vector of length (1:Size)
node_closeness() As Double()  'closeness of each node, vector of length (1:Size)
node_eigen() As Double()      'eigenvector centrality of each node, vector of length (1:Size)
node_Katz() As Double()      'Katz centrality of each node, vector of length (1:Size)
```

Read-Write properties
```  
node_pos(x() As Double)     'return or assign 2-D layout to nodes, array of size (1:Size, 1:2)
```

```Ruby
Print_Edges() as Variant
```
**Desc:** Return an array of size 3 x n_edge which is used to show edges in Excel on scatter chart.

```Ruby
ForceDirectedLayout(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 500)
```
**Desc:** Force directed algorithm to optimize graph layout, which is accessed by 'node_pos'.  Using quadratic repulsive force between nodes and log-linear spring.

**Input:**
- c1, spring constrant
- c2, spring length
- c3, repulsive force constant
- iter_max, maximum number of iterations allowed.

```Ruby
ForceDirectedLayout_BarnesHut(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 500)
```
**Desc:** Force directed algorithm accelerated with Barnes-Hut algroithm using quad tree.

**Input:** Same as `ForceDirectedLayout`

```Ruby
ForceDirected_MultiLevel(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 300)
```
**Desc:** Force directed algorithm accelerated with Barnes-Hut algroithm and multilevel

**Input:** Same as `ForceDirectedLayout`

```Ruby
Find_degree()
```
**Desc:** Calculate degree of each node which can then be accessed by 'node_degree'

```Ruby
Find_degree_wgt()
```
**Desc:** Calculate weighted degree of each node which can then be accessed by 'node_degree_wgt'. Weight is 1/distance to adjacent node.

```Ruby
Find_closeness()
```
**Desc:** Calculate closenss of each node which can then be accessed by 'node_closeness'. Using Dijkstra's Algorithm.

```Ruby
Find_Eigen(Optional iter_max As Long = 10000, Optional tolerance As Double = 0.0000000001)
```
**Desc:** Calculate eigen-centrality of each node which can then be accessed by 'node_eigen'. Using power iteration.

**Input:** iter_max and tolerance are the maximium iteration of tolerance in using power iteration.

```Ruby
Find_Katz(Optional alpha As Double = 1, Optional beta As Double = 1, _
        Optional iter_max As Long = 10000, Optional tol As Double = 0.0000000001)
```
**Desc:** Calculate Katz-centrality of each node which can then be accessed by 'node_Katz'. Using power iteration.

**Input:** Find c that satisfies c=alpha A c+beta, iter_max and tolerance are the maximium iteration of tolerance in using power iteration. c is normalized to ||c||=1

---
### mOutliers.bas
---
This module is used to calculate outlierscores in multidimensional data


#### Available methods:

```Ruby
MahalanobisDist(x() As Double) As Double()
```

**Desc:** Return Malanobis Distance of data x(1:N, 1:D). Ouput is a vector of length N.

**Input:** x(), a NxD array of data, where N is number of samples and D is dimension of data  

```Ruby
KthNeighborDist(x() As Double, Optional k As Long = 10, Optional usekdtree As Boolean = False) As Double()
```

**Desc:** Return k-th nearest neighor distance of x(1:N, 1:D). Ouput is a vector of length N. Uses Euclidean distance.

**Input:** 
- x(), a NxD array of data, where N is number of samples and D is dimension of data  
- k, number of nearest neighbors to consider
- usekdTree, use kd-Tree to speed up search when set to TRUE

```Ruby
LOF(x() As Double, Optional k As Long = 5) As Double()
```

**Desc:** Return local outlier factor of x(1:N, 1:D). Ouput is a vector of length N.

**Input:** 
- x(), a NxD array of data, where N is number of samples and D is dimension of data  
- k, number of nearest neighbors to consider

```Ruby
Influence_Iterate(x() As Double, Optional iterate_max As Long = 30, _
    Optional k_min As Long = 2, Optional k_max As Long = 10, Optional k_step As Long = 2) as Double()
```

**Desc:** Return influence of each data in x(1:N, 1:D). Ouput is a vector of length N. See '"Linear-Time Outlier Detection via Sensitivity", Mario Lucic, 2016 for reference.

**Input:** 
- x(), a NxD array of data, where N is number of samples and D is dimension of data
- iterate_max, number of realizations
- k_min & k_max & k_step, range of k-nearest neighbors to use in evaluation


---
### cSOM.cls
---
This class is used to create self-organizing map on a hexagonal grid


#### Available methods:

```Ruby
Init(L_x As Long, L_y As Long, dimension As Long)
```

**Desc:** Intialize grid to necesary size

**Input:**
- L_x & L_y, horizontal and vertical number of cells
- dimension, dimension of data

```Ruby
Read_Model(L_x As Long, L_y As Long, n_dimension As Long, node_w() As Double)
```

**Desc:** Intialize grid from previously saved node weights, which is an array of size (1:Lx, 1:Ly,1:D)

**Input:**
- L_x & L_y, horizontal and vertical number of cells
- dimension, dimension of data
- node_w, vectors of each node stored in an array of size (1:Lx, 1:Ly,1:D)

```Ruby
SOM_Hex_Train(x() As Double, _
        Optional iterate_max As Long = 5000, Optional learn_rate As Double = 0.1, _
        Optional batch_training As Boolean = False, Optional use_PCA As Boolean = True, _
        Optional random_sampling As Boolean = True)
```

**Desc:** Train SOM on raw data x

**Input:**
- x(), raw data as an array of size (1:N, 1:D), where N is number of samples and D is dimension
- iterate_max, maximum number of iterations
- learn_rate, learning rate
- batch_train, sequential training is performed if set to FALSE, otherwise batch mode is performed.
- use_PCA, PCA is used to initialize grid if set to TRUE, otherwise random initialization is performed.
- random_sampling, if batch_training is set to FALSE, then random sampling will feed in sample at random sequence.

```Ruby
Find_BMU(x() As Double, m As Long, n As Long, BMU_ED As Double)
```

**Desc:** Find best matching unit of x() on trained grid.

**Input:** x(), vector of length (1:D) representing one sample

**Output:**
- m & n , horizontal and vertical position on grid
- BMU_ED, distance from best matching unit.

```Ruby
Find_BMU_Batch(x() As Double, x_BMU() As Long, x_D2BMU() As Double)
```

**Desc:** Find best matching unit of x() on trained grid.

**Input:** x(), matrix of size (1:N, 1:D) representing N samples of D-dimensionl data

**Output:**
- x_BMU() , array of size (1:N,1:2),  horizontal and vertical positions on grid
- x_D2BMU(), vector of length 1:N, distance to best matching unit

**Output:**
- m & n , horizontal and vertical position on grid
- BMU_ED, distance from best matching unit.

```Ruby
Get_Node_Labels(x_name As Variant, node_label() As String)
```

**Desc:** for each node on the grad, stitch names of its members and output as label for that node.

**Input:** x_name(), vector of length (1:N) holding the name of each data.

**Output:** node_label(), array of size (1:L_x, 1:L_y), node lables of each node on the grid

**Read-Only Proeprties:**
```Ruby
quant_err() As Double()   'returns quatization error at each iteration
x_BMU() As Long()         'best matching unit of each data in the training set, array of size (1:N, 1:2)
x_D2BMU() As Double()     'distance to best matching unit of each data in the training set, vector of length (1:N)
wgts() As Double()        'feature vector of each node, array of size (1:L_x, 1:L_y, 1:D)
wgt(d as long) As Double() 'feature vector of each node in dimension d, array of size (1:L_x, 1:L_y)
wgts_norm() As Double()   'euclidean norm of feature vector of each node, array of size (1:L_x, 1:L_y)
UMatrix() As Double()     'avg distance of a node to its neighbours, array of size (1:L_x, 1:L_y).
```

```Ruby
Print_Network_results(vRng As Range)
```

**Desc:** Output network parameters to an Excel range
- 1st column: index
- 2nd column: row index of node
- 3rd column: column index of node
- 4th column: norm of node's vector
- 5th column: average distance to neighbor nodes
- 6th column: node vector


```Ruby
Print_All_Dimensions(mysht As Worksheet, Optional cht_width As Long = 280, Optional cht_height As Long = 280, Optional markersize As Long = 17, Optional write_labels As Boolean = False, Optional node_labels As Variant, Optional factor_names As Variant)
```

**Desc:** Print SOM charts on selected worksheet

**Input:**
- mysht, destination worksheet to print charts
- cht_width & cht_height, dimension of charts in points
- markersize, size of nodes in points
- write_labels, if TRUE then labels are displayed on charts.
- node_labels, if write_labels is set to TRUE then this will be used as node_labels
- factor_names, name of each dimension to be shown on top of each chart
